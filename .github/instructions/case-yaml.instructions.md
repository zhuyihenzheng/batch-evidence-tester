---
applyTo: "cases/**/*.yaml,cases/**/*.yml,templates/**/*.yaml,demo/**/cases/**/*.yaml,config/**/*.yaml"
---
用例 YAML / settings 的键一览与判定规则见 `docs/ai/case-schema.md`（未知键是硬错误）。
第 1 行有 `# generated-by: autotest gen` 的文件是生成物：改来源 md 后 `gen --overwrite`，不要直接改。
`enabled: false` 是未レビュー的印，不要擅自改成 `true`。注释用日文写「为什么」。
