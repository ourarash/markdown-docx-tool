# Mermaid Showcase

This sample demonstrates how Mermaid code fences can be rendered into the generated Word document as normal images.

## Simple Flowchart

```{.mermaid width=4.2in}
flowchart TD
  Draft[Write notes] --> Convert[Convert to DOCX]
  Convert --> Review[Review in Word]
```

## Styled Conversion Flow

```{.mermaid width=3.75in}
%%{init: {
  "theme": "base",
  "themeVariables": {
    "lineColor": "#3D5A80",
    "fontFamily": "Aptos, Segoe UI, sans-serif"
  }
}}%%
flowchart TD
  A([Draft markdown]):::source --> B{{Render Mermaid}}:::process
  B --> C[/Run Pandoc/]:::process
  C --> D([Review DOCX in Word]):::review

  classDef source fill:#E3F4E8,stroke:#2C7A5B,color:#123524,stroke-width:2px;
  classDef process fill:#FFF3D6,stroke:#C98A16,color:#5B3B00,stroke-width:2px;
  classDef review fill:#E6EEF8,stroke:#3D5A80,color:#1E314C,stroke-width:2px;
  linkStyle default stroke:#3D5A80,stroke-width:2px;
```

## Sequence Diagram

```{.mermaid width=5.25in}
%%{init: {
  "theme": "base",
  "themeVariables": {
    "actorBkg": "#E6EEF8",
    "actorBorder": "#3D5A80",
    "actorTextColor": "#1E314C",
    "signalColor": "#2C7A5B",
    "signalTextColor": "#214E3B",
    "noteBkgColor": "#FFF3D6",
    "noteBorderColor": "#C98A16",
    "noteTextColor": "#5B3B00",
    "labelBoxBkgColor": "#F4F7FB",
    "labelBoxBorderColor": "#9DB4D3",
    "labelTextColor": "#1E314C",
    "loopTextColor": "#1E314C",
    "fontFamily": "Aptos, Segoe UI, sans-serif"
  }
}}%%
sequenceDiagram
  autonumber
  participant Author
  participant Tool
  participant Word
  Author->>Tool: Convert report.md
  Tool->>Tool: Render Mermaid fences
  Tool->>Word: Write report.docx
  Note over Tool,Word: Mermaid diagrams are embedded<br/>as images in the DOCX.
  Word-->>Author: Show finished document
```
