# archive —— 已停用的历史实现

本目录保存 v1.0 时代的单体实现，**仅作历史参考，不再维护、不应运行**。

| 文件 | 说明 | 现状 |
|---|---|---|
| `rtf_merge_legacy.py` | v1.0 单体合并实现（1757 行，含 Word COM 后处理） | 已被 `rtfmerge/` 包替代 |
| `rtf_merge_gui_legacy.py` | v1.0 单体 GUI（332 行） | 已被 `rtfmerge/gui.py` 替代 |

- 现行入口：`rtf_merge.py`（CLI）、`rtf_merge_gui.py`（GUI）
- 现行实现：`rtfmerge/{merge_core,word_post,docx_xml,rtf_lint,cli,gui}.py`
- 现行后处理脚本：`rtfmerge/ps/{step1_convert,step2_find_empty,step3_recheck}.ps1`

这两个旧文件里仍留着 `if __name__ == "__main__":` 入口，但**直接运行它们会得到与当前交付不一致的结果**；
如需对比历史行为请仅在临时目录、用副本运行。

迁移时间：2026-11-06（v1.1.2，第三批收口）。
