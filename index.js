// cardo-status 的 Host 半：把包内 skill 注册到 ctx.skills。
// 依据《dsh 插件开发标准（现行版）》5.1、6.1。
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const name = 'cardo-status'
export const inject = ['skills']

const skillDir = join(dirname(fileURLToPath(import.meta.url)), 'skills', 'cardo-status')

function skillBody(raw) {
  // 去掉 YAML frontmatter：模型侧只接正文
  return raw.replace(/^---\r?\n[\s\S]*?\r?\n---\r?\n/, '')
}

export function apply(ctx) {
  ctx.skills.register({
    name: 'cardo-status',
    description: "Cardo 现状分析流水线：材料归集、齐套检查、事实抽取与《现状分析报告》成文。",
    whenToUse: "用户要基于现状底稿与项目资料产出《现状分析报告》《输入材料补齐确认单》《数据缺口清单》时使用。",
    source: 'custom',
    content: skillBody(readFileSync(join(skillDir, 'SKILL.md'), 'utf8')),
    resourceBase: { kind: 'directory', path: skillDir },
  })
}
