# Template And Layout Rules

## Established Layout

Use the accepted 条材厂-style sequence unless the supplied template differs:

1. cover;
2. owner approval page;
3. contents;
4. 一、编制依据及工程概况;
5. 二、施工组织体系;
6. 三、施工预设目标及组织管理体系;
7. 四、主要施工工序及方法;
8. 五、技术质量标准;
9. 六、质量保证措施;
10. 七、安全环保措施;
11. 八、资源需求计划;
12. 九、应急预案;
13. 十、施工进度计划.

Keep the main document portrait and the detailed final schedule landscape. Preserve the template's approval frame, signature spaces, footer, and page-number format.

## Typography

- Default body face: `方正仿宋_GB2312`.
- Preserve the template point sizes. Do not make every run the same size.
- Keep cover title, contents title, chapter headings, body, captions, source notes, and tables as distinct size levels.
- Use bold for chapter and subsection headings according to the template.
- Keep body paragraphs justified with the template's first-line indent and line spacing.
- Use smaller but readable caption, source-note, and table text.

When changing font face, set ASCII, high ANSI, East Asian, and complex-script font attributes at run level. Do not overwrite paragraph text or inherited size values.

## Tables

- Use tables for approval, quantities, roles, inspections, hazards, resources, emergency contacts, and schedules.
- Use prose for reasoning, sequence, method, controls, interfaces, and acceptance logic.
- Keep tables full-width but not crowded; use stable column widths.
- Repeat header rows and prevent row splitting where possible.
- Apply restrained pale blue or gray header shading and thin gray/black borders.
- Avoid nested cards or decorative table layouts.

## Figures And Standard Excerpts

- Center figures and keep aspect ratio.
- Keep the figure, caption, and source note together.
- Use a readable width; do not stretch low-resolution images.
- Do not leave generated empty labels or placeholder boxes in diagrams.
- Avoid forcing every figure onto a new page if that produces a nearly empty preceding page.
- If a full-page standard excerpt is needed, place related interpretation immediately before or after it.

## Safe DOCX Editing

- Preserve the original template.
- Prefer run-level replacement for retained cover, approval, and contents text.
- Remove and rebuild only the intended body region when the old project content must be replaced.
- Do not use `p.text = ...` on retained formatted paragraphs.
- Do not apply a global paragraph normalization pass.
- If the target is open in Word/WPS, save to a new filename rather than overwriting.
