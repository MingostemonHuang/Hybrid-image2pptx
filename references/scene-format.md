# Scene format

`source` is an image path relative to the scene file (or absolute). `canvas` is its exact `[width,height]` in pixels. Use `width_in:13.333333,height_in:7.5` for widescreen. All geometry and font sizes below are source pixels. Elements are drawn in array order. An explicit `id` becomes the native object's selection-pane name.

```json
{
  "source":"source.png", "canvas":[1672,941],
  "elements":[
    {"id":"solver","type":"shape","box":[30,136,790,247],"preset":"round_rect","radius":0.04,"fill":{"gradient":["#FFFFFF","#F6F6F8"],"angle":90},"stroke":"#CCCCD4","stroke_pt":0.5},
    {"id":"solver-title","type":"text","box":[55,144,300,55],"text":"AI4Solver","size_px":46,"bold":true,"color":"#B5121B"},
    {"id":"matrix","type":"picture","box":[315,193,487,140],"source_box":[315,193,487,140],"exclude_boxes":[[575,273,124,40]]},
    {"id":"agent","type":"shape","box":[390,547,634,180],"fill":"#FFFFFF","stroke":"#B5121B"},
    {"id":"solver-agent","type":"connector","kind":"straight","points":[[188,383],[188,456]],"from":"solver","to":"agent","color":"#B5121B","width_pt":1.5,"arrow":true}
  ]
}
```

Shapes: `rect`, `round_rect`, `ellipse`, `arrow_up`, `arrow_down`, `parallelogram`. Optional `shadow:true`, `shadow_alpha:13000`, `shadow_color:"#888888"`. A fill is a color string, a gradient with at least two colors, or omitted for transparent.

Text: optional `font`, `align:left|center|right`, `valign:top|middle|bottom`, `line_spacing`, `italic`, `wrap`. Default font is Microsoft YaHei. Explicit newlines preserve line grouping. For single-line mixed styling, provide `runs:[{"text":"...","color":"#...","bold":true}]`; inherited style comes from the containing text element. Use source-measured sizes rather than inventing a uniform font size.

Equations: use `type:"equation"`, `latex`, `box`, `size_px`, optional `color`, `bold`, `align` and `valign`. For example:

```json
{"id":"linear-system","type":"equation","box":[592,276,103,33],"source_box":[592,276,103,33],"latex":"Ax=b","size_px":29,"color":"#111111"}
```

`latex` contains the expression without `$` or `\\[...\\]` delimiters; JSON backslashes must be escaped. The builder creates a native Office Math object using Cambria Math and embeds LaTeX in its alternative description. LaTeX stays in the scene and the equation alternative description; no separate `.tex` file is generated. A source formula crop, from `source_box` or defaulting to `box`, is retained in an exclusive compatibility fallback. It must contain the complete formula and no surrounding prose. PowerPoint displays the native equation; a viewer without the corresponding equation support may display the crop. The fallback is not a second overlay. Use an independent equation box beside prose for mixed text/math; the builder does not parse LaTeX within ordinary text runs. Mathematical notation must be transcribed and checked before building. Unsupported commands stop conversion instead of silently becoming a picture.

Pictures: crop from the original source via `source_box:[x,y,w,h]`. `box` is placement. Optional `exclude_boxes` are source-pixel rectangles to remove from that crop for independently rebuilt text/plates. The crop is retained as PNG with no generative resampling. Preserve source aspect ratio in placement.

Connectors: `kind:straight|elbow|curve`. Two points give an editable native connector; four give an editable native freeform cubic `[start,control1,control2,end]` in source pixels. Optional `from`/`to` IDs preserve semantic node associations in the scene. Exported connectors keep measured geometry and remain independently editable; they are not automatically attached to nodes in this version. A decorative rule uses `arrow:false` without node IDs.

No source screenshot skin, no text-bearing image crop in place of editable labels, and no synthetic data graphic.
