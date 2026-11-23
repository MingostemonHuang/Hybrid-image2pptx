---
name: hybrid-image2pptx
description: Reconstruct slide screenshots and architecture diagrams as editable PPTX, using native text, containers, nodes and connectors while retaining complex scientific artwork from the source image. Use when both visual fidelity and structural editability matter. Not for generating a new slide design or flattening a page into one picture.
---

# Hybrid Image to PPTX

Preserve the supplied page while making its logical structure editable. The editable boundary is chosen per element, not once for the whole slide.

## Reconstruction decisions

- Read the current source image visually. Preserve wording, model names, values, topology, relative geometry and stacking order. Do not infer hidden chart data or change the design.
- Rebuild text, containers, node frames, title bars and diagram arrows as native PowerPoint objects. Use native connectors for relationships. Give nodes stable IDs and retain endpoint associations in the scene. Preserve measured routes; the current builder exports independently editable connector paths without automatic node attachment.
- Recognize mathematical expressions separately from prose. Transcribe their exact meaning into LaTeX and export them as native PowerPoint Office Math equations. Check subscripts, superscripts, Greek letters, signs and matrix dimensions against the source. Keep an exclusive source-crop fallback for viewers that do not support native equations. Do not place a bitmap over the native formula.
- Retain logos, CFD fields, photos, detailed matrices and 3D objects as separate source-derived pictures. Prefer original crops or alpha masks over generative reconstruction. Do not substitute illustrative flow fields for measured source results.
- A crop may contain only its intended artwork. If it contains text or a plate that will be editable, exclude that region from the picture and rebuild it natively; never leave duplicate text hidden underneath an editable overlay.
- Gradients and simple shadows should remain native where representable. Complex textures can be independent pictures. Do not rasterize an entire framework to retain its shadow.
- Source extraction, masking and cropping are asset preparation, not a substitute for image generation when a user requests new AI artwork. Use the available image-generation tool only for missing/occluded pixels or an explicitly requested replacement, and disclose reconstructed regions.

## Build

Create a new output directory and keep the original untouched. Use one source-pixel coordinate system for scene geometry, text size and crops. Read [scene format](references/scene-format.md) before authoring the scene.

The included builder uses Python with python-pptx, Pillow, lxml, latex2mathml and mathml2omml. Prefer the host bundled Python runtime (discover it with load_workspace_dependencies when available), or install requirements.txt in a local virtual environment. Do not modify unrelated presentation skills or global packages.

```bash
python scripts/build_pptx.py scene.json output.pptx
python scripts/inspect_pptx.py output.pptx
```

The scene is reusable editing input. Native objects get stable names. The builder supports solid/gradient fills, native shadows, source crops with exclusion masks, native text, connectors and editable cubic freeform paths. It deliberately does not draw a bitmap slide.

For equations, read the equation fields in [scene format](references/scene-format.md). The builder converts LaTeX through MathML to OMML and embeds the LaTeX source in each equation object. No separate `.tex` file is produced. Fractions, indices, integrals and matrices use structured equation elements. If conversion is unsupported or recognition is uncertain, resolve the expression before claiming native editability; never silently substitute an image and call it an editable equation.

## Verify what matters

- Render the actual PPTX with the target application when available. Inspect the native export, not just a preview of the scene JSON. Use a compatible renderer otherwise and identify material font substitution.
- Compare source and export at the same size: text wrapping, node positions, arrow direction, artwork identity, crop boundaries, gradients and layer order. Fix the owning object rather than hiding the discrepancy under a screenshot.
- Inspect the PPTX object structure: containers must be shapes, arrows must be connectors, text must be text. Source pictures must be confined to the intended artwork.
- Inspect equation objects for `a14:m/m:oMath`, not merely visible formula-like text. Native equations and their source previews are alternative representations, so only one is shown. A fallback-only rendering verifies fidelity, not PowerPoint equation editing; report that distinction when relevant.
- Check that scene endpoint IDs reference existing nodes and every required relationship is present. State that moving a node requires manually adjusting connector endpoints in this version. Never claim automatic rerouting or a manual drag test unless it has been implemented and verified.
- Fix concrete failures and render again after a material change. Do not keep running unchanged checks after they pass.

## Delivery

Return the PPTX and briefly state the editable boundary. Do not generate or deliver a separate `.tex` file. Include a source/export comparison when requested. Describe remaining approximation in one concise limitations paragraph. Never label a full-page bitmap with a few text overlays as an editable reconstruction.
