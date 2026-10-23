---
title: 排空课表工具
emoji: 📅
colorFrom: indigo
colorTo: blue
sdk: docker
app_port: 7860
pinned: false
---

# 排空课表工具

上传各部门课表 PDF 压缩包，自动生成部门空课表和每周系办值班表。

- 一次可传多个 zip，按部门文件夹自动识别
- 选择周次范围（每次都要选，不记忆）
- 输出：空课表.zip（每部门一张）+ 值班表.zip（每周一张）
