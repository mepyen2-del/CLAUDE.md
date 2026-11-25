
## SOLIDWORKS API & VBA Resources

A curated collection of open-source SOLIDWORKS automation projects, VBA macros, API examples, and engineering tools.

The primary target environment is **SOLIDWORKS 2019**, with a focus on reusable VBA code and backward-compatible API implementations.

### 1. VBA Macro Libraries

| Project | Main Features | Repository |
|---|---|---|
| CodeStack | SOLIDWORKS API, VBA macros, Drawing, Assembly, Properties | https://github.com/xarial/codestack |
| SOLIDWORKS Macro Library | PDF, DXF, STEP, BOM, batch export | https://github.com/SandBasher/solidworks-automation |
| Custom Properties Manager | Excel-based batch property management | https://github.com/ramugopal92/solidworks-custom-properties-manager |
| AutoDrawing | Automatic drawing views, dimensions and scale | https://github.com/kilwizac/solidworks-automate-drawing |
| Assembly Export | Recursive assembly traversal and batch export | https://github.com/mmmlinux/solidworks-export-macro |

### 2. SOLIDWORKS API Code Examples

| API Category | Reference |
|---|---|
| Assembly Traversal | https://www.codestack.net/solidworks-api/document/assembly/components/traversing-tree/ |
| Component Management | https://www.codestack.net/solidworks-api/document/assembly/components/ |
| Drawing Automation | https://www.codestack.net/solidworks-api/document/drawing/ |
| Dimension Traversal | https://www.codestack.net/solidworks-api/document/dimensions/traverse-all/ |
| Drawing View Scaling | https://www.codestack.net/solidworks-api/document/drawing/views-size-based-scale/ |
| Custom Properties | https://www.codestack.net/solidworks-api/data-storage/custom-properties/ |
| Document Management | https://www.codestack.net/solidworks-api/application/documents/ |

### 3. AI & MCP Automation

| Project | Main Features | Repository |
|---|---|---|
| SW_MCP | Python COM automation and AI tool integration | https://github.com/d-veloping/SW_MCP |
| SolidWorks MCP | Natural language CAD automation | https://github.com/alisamsam/solidworks-mcp |
| SolidWorks Automation Skill | AI Skill, MCP, verification and execution tracing | https://github.com/wzyn20051216/solidworks-automation-skill |

### 4. .NET Development Frameworks

| Project | Main Features | Repository |
|---|---|---|
| xCAD.NET | SOLIDWORKS Add-ins, Property Manager Pages and Macro Features | https://github.com/xarial/xcad |
| xCAD Examples | SOLIDWORKS API implementation examples | https://github.com/xarial/xcad-examples |

### 5. Compatibility Policy

- Primary development target: SOLIDWORKS 2019.
- Prefer native VBA macros and documented SOLIDWORKS API methods.
- Newer API examples must be reviewed and adapted before integration.
- C#, VB.NET and Python projects are reference implementations, not directly executable VBA macros.
- Every integrated module must pass compilation and functional testing in SOLIDWORKS 2019.
- Preserve existing modules and avoid deleting existing drawing dimensions or model data.

**Important:** Repository compatibility claims do not replace actual testing in SOLIDWORKS 2019.
