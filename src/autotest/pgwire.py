# -*- coding: utf-8 -*-
"""PostgreSQL のフロントエンド/バックエンドプロトコル（v3）の最小クライアント。

標準ライブラリだけで書いてある。試験機は閉じたネットワークにあることが多く
（RDP のクリップボード経由でしか物を持ち込めない環境もある）、psycopg2 や
psqlODBC を後から入れる前提にできない。このファイルならテキストとして
プロジェクトと一緒に運べる。

偽 OK を作らないための方針:
  - Simple Query だけを使う。パラメータのバインドも拡張プロトコルも使わない。
  - 値の型変換はしない。NULL は None、それ以外はサーバが返した文字表現を
    そのまま str にする（psql で見える値と同じ）。変換層が無ければ、変換の
    誤り（丸め・桁落ち・タイムゾーンの付け替え）も起きない。
  - 文字表現がサーバの設定で揺れないよう、DateStyle / bytea_output /
    extra_float_digits / client_encoding を接続時に固定する。
  - 判断できないこと（未対応の認証方式、SSL を要求したのに断られた、
    想定外の応答）は例外にする。黙って平文に落としたり、推測で続けたりしない。
"""

import base64
import hashlib
import hmac
import os
import socket
import ssl
import stringprep
import struct
import unicodedata
from typing import Dict, List, Optional, Sequence, Tuple

PROTOCOL_VERSION = 196608          # 3.0
SSL_REQUEST_CODE = 80877103
SSLMODES = ("disable", "prefer", "require", "verify-ca", "verify-full")
DEFAULT_PORT = 5432

# 接続時に必ず固定するセッション設定。値の文字表現をサーバ側の既定値に左右させない
FIXED_PARAMS = (
    ("client_encoding", "UTF8"),
    # 出力形式だけを ISO にする（入力の日付順 MDY/DMY はサーバの既定のまま）
    ("DateStyle", "ISO"),
    # escape 形式だと同じバイト列でも表記が変わる
    ("bytea_output", "hex"),
    # 12 未満のサーバは既定のままだと float を 15 桁に丸めて返す（12 以降は最短の正確な表記）
    ("extra_float_digits", "3"),
)

_AUTH_NAMES = {2: "Kerberos V5", 6: "SCM 資格情報", 7: "GSSAPI", 8: "GSSAPI", 9: "SSPI"}


class PgError(Exception):
    """サーバが返したエラー（ErrorResponse）か、接続・通信の失敗。

    fields はサーバのエラー項目（C = SQLSTATE、M = 本文、D = 詳細、H = ヒント …）。
    接続や通信の失敗では空になる。
    """

    def __init__(self, message: str, fields: Optional[Dict[str, str]] = None) -> None:
        super().__init__(message)
        self.fields = dict(fields or {})

    @property
    def sqlstate(self) -> str:
        return self.fields.get("C", "")

    @property
    def server_message(self) -> str:
        return self.fields.get("M", "")


class Result(object):
    """1 文ぶんの結果。columns が None なら行を返さない文（INSERT など）。"""

    def __init__(self, columns: Optional[List[str]], type_oids: Optional[List[int]]) -> None:
        self.columns = columns
        self.type_oids = type_oids
        self.rows = []  # type: List[List[Optional[str]]]
        self.command = ""


class Connection(object):
    """PostgreSQL への 1 接続。execute() は Simple Query を 1 回送る。

    1 回の execute() に複数の文を書くと、サーバはそれを 1 つの暗黙の
    トランザクションとして実行する（途中で失敗したら全体が取り消される）。
    """

    def __init__(self, host: str, port: int = DEFAULT_PORT, user: str = "",
                 password: Optional[str] = None, database: str = "",
                 sslmode: str = "prefer", sslrootcert: Optional[str] = None,
                 connect_timeout: float = 15.0, io_timeout: Optional[float] = None,
                 params: Sequence[Tuple[str, str]] = ()) -> None:
        if sslmode not in SSLMODES:
            raise PgError("sslmode は %s のいずれかです（指定値: %r）" % (" / ".join(SSLMODES), sslmode))
        self.host = host
        self.port = int(port)
        self.user = user
        self.password = password
        self.database = database
        self.sslmode = sslmode
        self.sslrootcert = sslrootcert
        self.connect_timeout = connect_timeout
        self.io_timeout = io_timeout
        self.params = list(FIXED_PARAMS) + [(k, v) for k, v in params]

        self.transaction_status = "I"      # I: 待機 / T: トランザクション中 / E: 失敗したトランザクション中
        self.server_params = {}  # type: Dict[str, str]
        self.notices = []  # type: List[str]
        self.ssl_version = ""              # 使った TLS のバージョン（平文なら空）
        self.auth_method = ""
        self.backend_pid = 0
        self._sock = None
        self._rfile = None
        self._broken = False
        self._scram = None  # type: Optional[_Scram]
        self._connect()

    # ------------------------------------------------------------------
    def __enter__(self) -> "Connection":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def execute(self, sql: str) -> List[Result]:
        """SQL を Simple Query で送り、文ごとの結果を返す。エラーなら PgError。

        エラーでも ReadyForQuery まで読み切るので、接続はそのまま使い続けられる。
        """
        if self._sock is None or self._broken:
            raise PgError("この接続はもう使えません（前の通信で失敗しています）")
        if "\0" in sql:
            raise PgError("SQL に NUL 文字が含まれています")
        self._send(b"Q", _cstr(sql))

        results = []  # type: List[Result]
        current = None  # type: Optional[Result]
        error = None  # type: Optional[PgError]
        while True:
            kind, data = self._read_message()
            if kind == b"T":
                names, oids = _parse_row_description(data)
                current = Result(names, oids)
            elif kind == b"D":
                if current is None:
                    self._fail("列定義の前に行データを受信しました")
                current.rows.append(_parse_data_row(data))
            elif kind == b"C":
                if current is None:
                    current = Result(None, None)
                current.command = _text(data.rstrip(b"\0"))
                results.append(current)
                current = None
            elif kind == b"I":      # EmptyQueryResponse（空文字や ; だけの SQL）
                current = None
            elif kind == b"E":
                error = self._error(data)
                current = None
            elif kind == b"G":      # COPY FROM STDIN は流し込むデータを持っていない
                self._send(b"f", _cstr("autotest は COPY FROM STDIN に対応していません"))
            elif kind == b"H":      # COPY TO STDOUT の中身は読み捨てて失敗にする
                error = error or PgError("COPY TO STDOUT には対応していません（SELECT で取得してください）")
            elif kind in (b"d", b"c"):
                pass
            elif kind == b"Z":
                self.transaction_status = _text(data[:1])
                break
            else:
                self._handle_async(kind, data)
        if error is not None:
            raise error
        return results

    def close(self) -> None:
        sock, self._sock = self._sock, None
        if sock is None:
            return
        try:
            if not self._broken:
                sock.sendall(b"X\x00\x00\x00\x04")    # Terminate
        except (OSError, ssl.SSLError):
            pass
        finally:
            try:
                if self._rfile is not None:
                    self._rfile.close()
            finally:
                sock.close()

    # --- 接続 -----------------------------------------------------------
    def _connect(self) -> None:
        try:
            sock = socket.create_connection((self.host, self.port), timeout=self.connect_timeout)
        except socket.timeout:
            raise PgError("接続がタイムアウトしました（%s 秒）: %s:%d" % (self.connect_timeout, self.host, self.port))
        except OSError as exc:
            raise PgError("サーバに接続できません: %s:%d（%s）" % (self.host, self.port, exc))
        try:
            sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
            sock = self._negotiate_ssl(sock)
            self._sock = sock
            self._rfile = sock.makefile("rb")
            self._startup()
        except BaseException:
            self._broken = True
            self._sock = sock
            self.close()
            raise
        # 認証までは connect_timeout、その後の問い合わせは io_timeout で待つ
        sock.settimeout(self.io_timeout)

    def _negotiate_ssl(self, sock):
        if self.sslmode == "disable":
            return sock
        sock.sendall(struct.pack("!ii", 8, SSL_REQUEST_CODE))
        # 1 バイトだけ読む。余分に読むと、TLS より前に第三者が差し込んだデータを
        # 平文のまま受け入れてしまう
        answer = sock.recv(1)
        if answer == b"N":
            if self.sslmode in ("require", "verify-ca", "verify-full"):
                raise PgError("サーバが SSL を受け付けませんでした（sslmode=%s）。"
                              "サーバの ssl 設定を確認するか、暗号化しなくてよい環境なら"
                              " sslmode を prefer / disable にしてください" % self.sslmode)
            return sock
        if answer != b"S":
            raise PgError("SSL の交渉で想定外の応答がありました: %r" % answer)
        context = self._ssl_context()
        try:
            wrapped = context.wrap_socket(sock, server_hostname=self.host)
        except (ssl.SSLError, ssl.CertificateError, OSError) as exc:
            raise PgError("SSL を確立できませんでした（sslmode=%s）: %s" % (self.sslmode, exc))
        self.ssl_version = wrapped.version() or "TLS"
        return wrapped

    def _ssl_context(self) -> ssl.SSLContext:
        context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        if self.sslmode in ("prefer", "require"):
            # libpq と同じく、この 2 つは暗号化だけで証明書は検証しない
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE
            return context
        context.verify_mode = ssl.CERT_REQUIRED
        context.check_hostname = self.sslmode == "verify-full"
        if self.sslrootcert:
            try:
                context.load_verify_locations(cafile=self.sslrootcert)
            except (OSError, ssl.SSLError) as exc:
                raise PgError("sslrootcert を読み込めません: %s（%s）" % (self.sslrootcert, exc))
        else:
            context.load_default_certs()
        return context

    def _startup(self) -> None:
        body = struct.pack("!i", PROTOCOL_VERSION)
        for key, value in [("user", self.user), ("database", self.database)] + self.params:
            body += _cstr(key) + _cstr(str(value))
        body += b"\0"
        self._sendall(struct.pack("!i", len(body) + 4) + body)
        while True:
            kind, data = self._read_message()
            if kind == b"R":
                self._authenticate(data)
            elif kind == b"E":
                raise self._error(data)
            elif kind == b"Z":
                self.transaction_status = _text(data[:1])
                return
            else:
                self._handle_async(kind, data)

    def _authenticate(self, data: bytes) -> None:
        code = struct.unpack_from("!i", data)[0]
        if code == 0:
            if self._scram is not None and not self._scram.verified:
                # 署名の確認前に「成功」とだけ言うサーバは、本物か確かめられない
                self._fail("SCRAM 認証が完了する前に認証成功が返されました")
            self.auth_method = self.auth_method or "trust"
        elif code == 3:
            self.auth_method = "password"
            self._send(b"p", _cstr(self._need_password()))
        elif code == 5:
            self.auth_method = "md5"
            self._send(b"p", _cstr(md5_password(self._need_password(), self.user, data[4:8])))
        elif code == 10:
            mechanisms = [m for m in data[4:].split(b"\0") if m]
            if b"SCRAM-SHA-256" not in mechanisms:
                raise PgError("サーバが提示した SASL 方式に対応していません: %s"
                              % ", ".join(_text(m) for m in mechanisms))
            self.auth_method = "SCRAM-SHA-256"
            self._scram = _Scram(self._need_password())
            first = self._scram.client_first()
            self._send(b"p", _cstr("SCRAM-SHA-256") + struct.pack("!i", len(first)) + first)
        elif code == 11:
            if self._scram is None:
                self._fail("SASL の手順外のメッセージを受信しました")
            self._send(b"p", self._scram.client_final(data[4:]))
        elif code == 12:
            if self._scram is None:
                self._fail("SASL の手順外のメッセージを受信しました")
            self._scram.verify_server(data[4:])
        else:
            raise PgError("未対応の認証方式です: %s。サーバの pg_hba.conf で、この接続の METHOD を"
                          " scram-sha-256 か md5 にしてください" % _AUTH_NAMES.get(code, "コード %d" % code))

    def _need_password(self) -> str:
        if self.password is None:
            raise PgError("サーバがパスワードを要求しましたが、パスワードが設定されていません"
                          "（settings の database.password_env）")
        return self.password

    # --- 送受信 ---------------------------------------------------------
    def _send(self, kind: bytes, payload: bytes) -> None:
        self._sendall(kind + struct.pack("!i", len(payload) + 4) + payload)

    def _sendall(self, data: bytes) -> None:
        try:
            self._sock.sendall(data)
        except socket.timeout:
            self._broken = True
            raise PgError("サーバへの送信がタイムアウトしました")
        except (OSError, ssl.SSLError) as exc:
            self._broken = True
            raise PgError("サーバへの送信に失敗しました: %s" % exc)

    def _read_message(self) -> Tuple[bytes, bytes]:
        header = self._read_exact(5)
        length = struct.unpack("!i", header[1:])[0]
        if length < 4:
            self._fail("メッセージ長が不正です: %d" % length)
        return header[:1], self._read_exact(length - 4)

    def _read_exact(self, size: int) -> bytes:
        try:
            data = self._rfile.read(size) if size else b""
        except socket.timeout:
            self._broken = True
            raise PgError("サーバからの応答がありません（%s 秒）" % self._sock.gettimeout())
        except (OSError, ssl.SSLError) as exc:
            self._broken = True
            raise PgError("サーバとの通信に失敗しました: %s" % exc)
        if data is None or len(data) < size:
            self._broken = True
            raise PgError("サーバとの接続が切れました")
        return data

    def _handle_async(self, kind: bytes, data: bytes) -> None:
        if kind == b"S":            # ParameterStatus
            parts = data.split(b"\0")
            if len(parts) >= 2:
                self.server_params[_text(parts[0])] = _text(parts[1])
        elif kind == b"K":          # BackendKeyData
            self.backend_pid = struct.unpack_from("!i", data)[0]
        elif kind == b"N":          # NoticeResponse（WARNING / NOTICE）。証跡の調査用に残す
            self.notices.append(format_fields(parse_fields(data)))
        elif kind in (b"A", b"v"):  # LISTEN の通知 / プロトコル版の交渉（どちらも使わない）
            pass
        else:
            self._fail("想定外のメッセージを受信しました: %r" % kind)

    def _error(self, data: bytes) -> PgError:
        fields = parse_fields(data)
        return PgError(format_fields(fields), fields)

    def _fail(self, message: str) -> None:
        # 応答の読み筋を見失ったら接続ごと捨てる。続けると前の結果を次の問い合わせの結果と取り違える
        self._broken = True
        raise PgError("プロトコルエラー: %s" % message)


# =============================================================================
# 認証
# =============================================================================


def md5_password(password: str, user: str, salt: bytes) -> str:
    inner = hashlib.md5((password + user).encode("utf-8")).hexdigest()
    return "md5" + hashlib.md5(inner.encode("ascii") + salt).hexdigest()


class _Scram(object):
    """SCRAM-SHA-256（RFC 5802 / 7677）のクライアント側。チャネルバインディングは使わない。

    PostgreSQL はユーザー名を起動メッセージのものを使うため、SCRAM の n= は空で送る。
    """

    def __init__(self, password: str, nonce: Optional[str] = None, user: str = "") -> None:
        self.password = saslprep(password).encode("utf-8")
        self.nonce = nonce or base64.b64encode(os.urandom(18)).decode("ascii")
        self.first_bare = "n=%s,r=%s" % (user, self.nonce)
        self.verified = False
        self._server_signature = b""

    def client_first(self) -> bytes:
        return ("n,," + self.first_bare).encode("utf-8")

    def client_final(self, server_first: bytes) -> bytes:
        text = server_first.decode("utf-8")
        attrs = _scram_attrs(text)
        nonce = attrs.get("r", "")
        if not nonce.startswith(self.nonce) or len(nonce) <= len(self.nonce):
            raise PgError("SCRAM: サーバの nonce が不正です")
        if "m" in attrs:
            raise PgError("SCRAM: 未対応の拡張が要求されました")
        try:
            salt = base64.b64decode(attrs["s"])
            iterations = int(attrs["i"])
        except (KeyError, ValueError, TypeError):
            raise PgError("SCRAM: サーバの応答（salt / 反復回数）が不正です")
        if iterations < 1:
            raise PgError("SCRAM: 反復回数が不正です: %d" % iterations)
        salted = hashlib.pbkdf2_hmac("sha256", self.password, salt, iterations)
        client_key = hmac.new(salted, b"Client Key", hashlib.sha256).digest()
        stored_key = hashlib.sha256(client_key).digest()
        without_proof = "c=biws,r=" + nonce
        auth_message = ",".join((self.first_bare, text, without_proof)).encode("utf-8")
        signature = hmac.new(stored_key, auth_message, hashlib.sha256).digest()
        proof = bytes(a ^ b for a, b in zip(client_key, signature))
        server_key = hmac.new(salted, b"Server Key", hashlib.sha256).digest()
        self._server_signature = hmac.new(server_key, auth_message, hashlib.sha256).digest()
        return (without_proof + ",p=" + base64.b64encode(proof).decode("ascii")).encode("ascii")

    def verify_server(self, server_final: bytes) -> None:
        attrs = _scram_attrs(server_final.decode("utf-8"))
        if "e" in attrs:
            raise PgError("SCRAM: サーバが認証エラーを返しました: %s" % attrs["e"])
        try:
            got = base64.b64decode(attrs.get("v", ""))
        except (ValueError, TypeError):
            got = b""
        if not self._server_signature or not hmac.compare_digest(got, self._server_signature):
            raise PgError("SCRAM: サーバの署名が一致しません（接続先が本物のサーバではない可能性）")
        self.verified = True


def _scram_attrs(text: str) -> Dict[str, str]:
    attrs = {}
    for part in text.split(","):
        if len(part) < 2 or part[1] != "=":
            raise PgError("SCRAM: サーバの応答を解釈できません: %r" % text)
        attrs[part[0]] = part[2:]
    return attrs


def saslprep(text: str) -> str:
    """RFC 4013 の SASLprep。サーバも保存時に同じ正規化をするので合わせる。

    正規化できない文字列（禁止文字を含む等）は、サーバと同じく元のまま使う。
    """
    if all(ord(ch) < 0x80 for ch in text):
        return text
    mapped = "".join(" " if stringprep.in_table_c12(ch) else ch
                     for ch in text if not stringprep.in_table_b1(ch))
    normalized = unicodedata.normalize("NFKC", mapped)
    if not normalized:
        return text
    prohibited = (stringprep.in_table_a1, stringprep.in_table_c12, stringprep.in_table_c21_c22,
                  stringprep.in_table_c3, stringprep.in_table_c4, stringprep.in_table_c5,
                  stringprep.in_table_c6, stringprep.in_table_c7, stringprep.in_table_c8,
                  stringprep.in_table_c9)
    if any(check(ch) for ch in normalized for check in prohibited):
        return text
    if any(stringprep.in_table_d1(ch) for ch in normalized):
        if any(stringprep.in_table_d2(ch) for ch in normalized):
            return text
        if not (stringprep.in_table_d1(normalized[0]) and stringprep.in_table_d1(normalized[-1])):
            return text
    return normalized


# =============================================================================
# メッセージの解釈
# =============================================================================


def _cstr(text: str) -> bytes:
    return text.encode("utf-8") + b"\0"


def _text(data: bytes) -> str:
    try:
        return data.decode("utf-8")
    except UnicodeDecodeError:
        # 置換文字で読み替えると値が変わったまま比較に進む。読めないものは読めないと言う
        raise PgError("UTF-8 として読めないデータを受信しました（DB の文字コードが SQL_ASCII 等の可能性）")


def _parse_row_description(data: bytes) -> Tuple[List[str], List[int]]:
    count = struct.unpack_from("!h", data, 0)[0]
    pos = 2
    names, oids = [], []
    for _ in range(count):
        end = data.index(b"\0", pos)
        names.append(_text(data[pos:end]))
        _table, _attnum, type_oid, _typlen, _typmod, _fmt = struct.unpack_from("!IhIhih", data, end + 1)
        pos = end + 19
        oids.append(type_oid)
    return names, oids


def _parse_data_row(data: bytes) -> List[Optional[str]]:
    count = struct.unpack_from("!h", data, 0)[0]
    pos = 2
    row = []  # type: List[Optional[str]]
    for _ in range(count):
        size = struct.unpack_from("!i", data, pos)[0]
        pos += 4
        if size < 0:
            row.append(None)        # 長さ -1 が NULL。長さ 0 の空文字とは別物
        else:
            row.append(_text(data[pos:pos + size]))
            pos += size
    return row


def parse_fields(data: bytes) -> Dict[str, str]:
    """ErrorResponse / NoticeResponse の項目（1 文字のコード + 文字列）を辞書にする。"""
    fields = {}
    pos = 0
    while pos < len(data) and data[pos:pos + 1] != b"\0":
        code = data[pos:pos + 1].decode("ascii", "replace")
        end = data.index(b"\0", pos + 1)
        fields[code] = data[pos + 1:end].decode("utf-8", "replace")
        pos = end + 1
    return fields


def format_fields(fields: Dict[str, str]) -> str:
    severity = fields.get("S") or fields.get("V") or "ERROR"
    text = "%s: %s" % (severity, fields.get("M", ""))
    if fields.get("C"):
        text += "（SQLSTATE %s）" % fields["C"]
    for code, label in (("D", "詳細"), ("H", "ヒント"), ("W", "場所")):
        if fields.get(code):
            text += "\n  %s: %s" % (label, fields[code])
    if fields.get("P"):
        text += "\n  位置: SQL の %s 文字目" % fields["P"]
    return text
