$ErrorActionPreference = "Stop"

$workDir = "C:\Users\namma\.claude\api_spec_extract_work"
$outDir = Join-Path $workDir "output"
if (!(Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

function New-WordApp {
    try {
        $word = New-Object -ComObject Word.Application
    } catch {
        $wordKey = "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency"
        foreach ($child in @("StartupItems", "DocumentRecovery", "DisabledItems")) {
            $path = Join-Path $wordKey $child
            if (Test-Path -LiteralPath $path) {
                Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        $exe = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
        if (Test-Path -LiteralPath $exe) {
            Start-Process -FilePath $exe -WindowStyle Hidden | Out-Null
            Start-Sleep -Seconds 3
        }
        $word = New-Object -ComObject Word.Application
    }
    $word.Visible = $false
    $word.DisplayAlerts = 0
    return $word
}

function Add-Para {
    param(
        [object]$Doc,
        [object]$Sel,
        [string]$Text,
        [string]$Kind = "Normal"
    )
    switch ($Kind) {
        "Title" {
            $Sel.Style = $Doc.Styles.Item([int]-63)
            $Sel.Font.Name = "Aptos Display"
            $Sel.Font.Size = [single]20
            $Sel.Font.Bold = [int]1
            $Sel.ParagraphFormat.SpaceAfter = [single]8
        }
        "Subtitle" {
            $Sel.Style = $Doc.Styles.Item([int]-1)
            $Sel.Font.Name = "Aptos"
            $Sel.Font.Size = [single]10.5
            $Sel.Font.Bold = [int]0
            $Sel.Font.Italic = [int]1
            $Sel.ParagraphFormat.SpaceAfter = [single]14
        }
        "H1" {
            $Sel.Style = $Doc.Styles.Item([int]-2)
            $Sel.Font.Name = "Aptos Display"
            $Sel.Font.Size = [single]14
            $Sel.Font.Bold = [int]1
            $Sel.Font.Italic = [int]0
            $Sel.ParagraphFormat.SpaceBefore = [single]10
            $Sel.ParagraphFormat.SpaceAfter = [single]4
        }
        "H2" {
            $Sel.Style = $Doc.Styles.Item([int]-3)
            $Sel.Font.Name = "Aptos"
            $Sel.Font.Size = [single]11.5
            $Sel.Font.Bold = [int]1
            $Sel.Font.Italic = [int]0
            $Sel.ParagraphFormat.SpaceBefore = [single]8
            $Sel.ParagraphFormat.SpaceAfter = [single]3
        }
        "Bullet" {
            $Sel.Style = $Doc.Styles.Item([int]-49)
            $Sel.Font.Name = "Aptos"
            $Sel.Font.Size = [single]9.5
            $Sel.Font.Bold = [int]0
            $Sel.Font.Italic = [int]0
            $Sel.ParagraphFormat.SpaceAfter = [single]2
        }
        "Code" {
            $Sel.Style = $Doc.Styles.Item([int]-1)
            $Sel.Font.Name = "Consolas"
            $Sel.Font.Size = [single]8.5
            $Sel.Font.Bold = [int]0
            $Sel.Font.Italic = [int]0
            $Sel.ParagraphFormat.LeftIndent = [single]18
            $Sel.ParagraphFormat.SpaceAfter = [single]2
        }
        default {
            $Sel.Style = $Doc.Styles.Item([int]-1)
            $Sel.Font.Name = "Aptos"
            $Sel.Font.Size = [single]9.5
            $Sel.Font.Bold = [int]0
            $Sel.Font.Italic = [int]0
            $Sel.ParagraphFormat.LeftIndent = [single]0
            $Sel.ParagraphFormat.SpaceAfter = [single]4
        }
    }
    $Sel.TypeText($Text)
    $Sel.TypeParagraph()
    $Sel.ParagraphFormat.LeftIndent = [single]0
}

function Build-Doc {
    param(
        [object]$Word,
        [string]$Title,
        [string]$Subtitle,
        [array]$Blocks,
        [string]$BaseName
    )
    $docx = Join-Path $outDir ($BaseName + ".docx")
    $pdf = Join-Path $outDir ($BaseName + ".pdf")

    $doc = $Word.Documents.Add()
    try {
        $doc.PageSetup.TopMargin = $Word.InchesToPoints(0.65)
        $doc.PageSetup.BottomMargin = $Word.InchesToPoints(0.65)
        $doc.PageSetup.LeftMargin = $Word.InchesToPoints(0.7)
        $doc.PageSetup.RightMargin = $Word.InchesToPoints(0.7)

        $doc.Styles.Item([int]-2).Font.Color = [int]10485760
        $doc.Styles.Item([int]-3).Font.Color = [int]52479

        $sel = $Word.Selection
        Add-Para $doc $sel $Title "Title"
        Add-Para $doc $sel $Subtitle "Subtitle"

        foreach ($b in $Blocks) {
            Add-Para $doc $sel $b.Text $b.Kind
        }

        $sel.Style = $doc.Styles.Item([int]-1)
        $sel.Font.Bold = [int]0
        $sel.Font.Italic = [int]0
        $sel.Font.Name = "Aptos"

        $footer = $doc.Sections.Item(1).Footers.Item(1).Range
        $footer.Text = "API specification extraction | prepared 2026-06-14 | "
        $footer.Font.Name = "Aptos"
        $footer.Font.Size = [single]8
        $footer.ParagraphFormat.Alignment = 1
        $doc.Sections.Item(1).Footers.Item(1).PageNumbers.Add() | Out-Null

        $doc.SaveAs2([string]$docx)
        $doc.ExportAsFixedFormat($pdf, 17)
    } finally {
        $doc.Close([ref]$false)
    }

    return [pscustomobject]@{ Docx = $docx; Pdf = $pdf }
}

$dateNote = "Prepared in English on 2026-06-14 from publicly reachable official pages, public mirrors, and public SDK source code. This is a paraphrased extraction, not a reproduction of vendor manuals."

$avevaNet = @(
    @{Kind="H1"; Text="Source Access Status"},
    @{Kind="Normal"; Text="The exact current AVEVA Everything3D .NET Customisation Guide appears to be restricted or only partially public. The extraction below is based on AVEVA public E3D help pages, AVEVA documentation snippets, and a publicly mirrored AVEVA .NET Customisation User Guide for the 12 Series. Treat names and namespaces as version-sensitive."},
    @{Kind="H1"; Text="API Surface Overview"},
    @{Kind="Bullet"; Text="Two principal extension paths are exposed: Common Application Framework add-ins written in .NET, and PML.NET objects that let PML instantiate and call selected .NET classes."},
    @{Kind="Bullet"; Text="CAF add-ins run inside the AVEVA host process and share the same database session, command context, UI services, and event loop as the module in which they are loaded."},
    @{Kind="Bullet"; Text="PML.NET is a bridge for PML-to-.NET invocation. It uses PML proxy objects backed by .NET classes. Only callable classes that satisfy the PML.NET rules should be exposed."},
    @{Kind="Bullet"; Text="Public E3D help shows namespace drift: E3D 1.1 file-browser examples use Aveva.Pdms.Presentation, while E3D 3.1 examples use Aveva.Core.Presentation. Verify the target product release before compiling."},
    @{Kind="H1"; Text="Main Assemblies and Namespaces"},
    @{Kind="Bullet"; Text="Application framework: Aveva.ApplicationFramework.dll and Aveva.ApplicationFramework.Presentation.dll provide host integration, commands, windows, command bars, resource loading, and presentation services."},
    @{Kind="Bullet"; Text="Database API: legacy examples use Aveva.Pdms.Database; later E3D examples frequently use Aveva.Core.Database. Key object families include DbElement, DbElementType, DbAttribute, DB/MDB/project/session objects, filters, iterators, and events."},
    @{Kind="Bullet"; Text="Other AVEVA API families named by the guide include geometry, shared utilities, graphics, design, piping, admin, clash, data-management, and presentation controls. Availability depends on installed modules and licenses."},
    @{Kind="H1"; Text="CAF Add-In Specification"},
    @{Kind="Bullet"; Text="A CAF add-in is packaged as a .NET assembly and configured for loading by the target AVEVA module. The guide describes the add-in interface, startup performance, window management, command classes, command events, resource manager, tracing, exception handling, and thread safety."},
    @{Kind="Bullet"; Text="UI exposure is normally through UIC customisation files and the CommandBarManager. A module customisation XML file lists UIC files loaded at startup. The CAF_UIC_PATH environment variable can redirect development copies without editing the install folder."},
    @{Kind="Bullet"; Text="Add-in commands should isolate host event-loop entry points. General .NET exceptions should be caught by add-in code and translated or handled; allowing arbitrary exceptions to escape can destabilize the AVEVA host."},
    @{Kind="Bullet"; Text="Threading must be conservative: the host core, CAF, and database layer are not generally multi-thread-safe. Only one thread should call into host core services at a time."},
    @{Kind="H1"; Text="Database API Specification"},
    @{Kind="Bullet"; Text="DbElementType identifies element kinds. Standard instances are exposed through DbElementTypeInstance; dynamic or user-defined types can be resolved by name or hash. Metadata includes name, description, base type, valid attributes, valid owners/members, and database type coverage."},
    @{Kind="Bullet"; Text="DbAttribute identifies an attribute definition, not a value. Standard instances are exposed through DbAttributeInstance; user-defined attributes can be resolved by name or hash. Metadata includes type, units, category, size, allowed values/ranges, UDA flag, pseudo-attribute flag, and qualifier rules."},
    @{Kind="Bullet"; Text="DbAttributeType covers the core scalar and geometric categories: integer, double, Boolean, string, element reference, direction, position, and orientation."},
    @{Kind="Bullet"; Text="DbElement is the main access object for database instances. Typical operations include locating an element by reference/name, navigating owner/member/next/previous links, reading attributes by typed getters, checking valid attributes, setting modifiable attributes, creating, copying, deleting, reordering, and changing type where legal."},
    @{Kind="Bullet"; Text="Navigation pseudo-attributes include member lists, owner lists, sequence position, hierarchy depth, all elements by type, and secondary hierarchy expansion. Secondary hierarchies require pseudo-attributes such as SMEMB/SEXPND rather than ordinary first-member navigation."},
    @{Kind="Bullet"; Text="Database modification is guarded by access and legality checks. Pseudo-attributes such as DAC and MOD families expose whether creation, deletion, copying, claiming, exporting, issuing, and attribute modification are allowed and what error would occur."},
    @{Kind="Bullet"; Text="Expressions are represented by DbExpression. Expressions can be parsed from text, stored with rules/catalogue parameterisation, converted back to text, and evaluated against a DbElement with a result type such as double, string, Boolean, element, position, direction, or orientation."},
    @{Kind="Bullet"; Text="Units are explicit. For numeric expression evaluation, callers should pass the expected unit category where ambiguity exists. Raw database double values and formatted user values are different concerns."},
    @{Kind="Bullet"; Text="Events can be subscribed from C# for general database changes and DB/MDB lifecycle activity. Public community examples identify DBEvents in Aveva.Core.Database for project/MDB open-close, savework/getwork, undo/redo, refresh, drop, and user-change notifications."},
    @{Kind="H1"; Text="PML.NET Bridge Specification"},
    @{Kind="Bullet"; Text="PML loads a .NET class by proxy class name plus the active namespace. Namespace names are case-insensitive in the PML layer, but partial namespace concatenation is not supported."},
    @{Kind="Bullet"; Text="Only .NET classes marked and structured as PML.NET callable can be called from PML. Object names are case-insensitive and restricted by PML naming rules."},
    @{Kind="Bullet"; Text="PML can instantiate .NET objects, query variables, and list callable methods. PML can pass primitive values, strings, arrays, database references, and existing PML.NET proxy instances."},
    @{Kind="Bullet"; Text="PML system/user objects such as direction and orientation are not generally passed directly as rich objects. Pass database references or supported scalar/array data instead."},
    @{Kind="Bullet"; Text=".NET cannot directly call arbitrary PML objects; callbacks/events are the supported route back into PML."},
    @{Kind="Code"; Text="PML pattern: import 'AssemblyOrObjectName'; using namespace 'Aveva.Core.Presentation'; !obj = object ClassName(args); q methods !obj"},
    @{Kind="H1"; Text="AVEVA Grid Control and Presentation Controls"},
    @{Kind="Bullet"; Text="The AVEVA Grid Control can bind database element/attribute data, non-database tabular arrays, or Excel/CSV-like file sources. It supports layout save/load, grouping, sorting, filters, fixed rows/headers, editable cells, row operations, column operations, cell formatting, drag/drop, print preview, and export to Excel."},
    @{Kind="Bullet"; Text="Editable grid cells do not automatically write back to Dabacon. The caller must handle BeforeCellUpdate or related events and explicitly synchronize database attributes, commonly through DoDabaconCellUpdate when applicable."},
    @{Kind="Bullet"; Text="The PML File Browser wrapper is exposed as PMLFileBrowser. Its show method accepts default directory, seed filename, title, exists flag, filter string, and selected filter index; selected file is queried through file()."},
    @{Kind="Code"; Text="PML File Browser: !browser.show(directory, seed, title, exists, filter, index); q var !browser.file()"},
    @{Kind="H1"; Text="Practical Integration Notes"},
    @{Kind="Bullet"; Text="Version-check every namespace and assembly reference before implementation. E3D 3.x moved public examples from Aveva.Pdms.* toward Aveva.Core.* in at least some presentation namespaces."},
    @{Kind="Bullet"; Text="Keep customization packages isolated from vendor installation folders where possible. Use project-level or environment-variable configuration for UIC/add-in load paths."},
    @{Kind="Bullet"; Text="Treat database write operations as transactional business logic: check legality, claim state, access control, and user-visible error text before changing model data."},
    @{Kind="Bullet"; Text="For marine/plant PLM integration, use add-ins or local agents to mediate between E3D authoring and an external control plane. Do not assume that a remote REST service can call E3D internals directly without an in-process or desktop-side component."},
    @{Kind="H1"; Text="Sources Used"},
    @{Kind="Bullet"; Text="AVEVA E3D 1.1 .NET Customisation User Guide public page: https://help.aveva.com/AVEVA_Everything3D/1.1/NCUG/NCUG8.html"},
    @{Kind="Bullet"; Text="AVEVA E3D 3.1 .NET Customisation User Guide public page: https://help.aveva.com/AVEVA_Everything3D/3.1/NCUG/NCUG8.html"},
    @{Kind="Bullet"; Text="AVEVA PML.NET documentation page/snippet: https://docs.aveva.com/bundle/engineering/page/962704.html"},
    @{Kind="Bullet"; Text="AVEVA Software Customisation Guide public page for CMSYS/.NET command manager: https://help.aveva.com/AVEVA_Everything3D/1.1/SOFTCG/SOFTCG29.html"},
    @{Kind="Bullet"; Text="Public mirror of AVEVA .NET Customisation User Guide, used only for paraphrased extraction because official full guide access was restricted: https://pdfcoffee.com/net-customisation-user-guidepdf-pdf-free.html"}
)

$avevaPml = @(
    @{Kind="H1"; Text="Source Access Status"},
    @{Kind="Normal"; Text="The official AVEVA PML reference is partially available through AVEVA public help and AVEVA documentation pages. The detailed object/method tables are largely restricted, so this extraction combines AVEVA public pages with a public mirror of the Software Customisation Reference Manual. It is written as an engineering index, not as a replacement for the licensed manual."},
    @{Kind="H1"; Text="PML Language Model"},
    @{Kind="Bullet"; Text="PML has two generations. PML 1 remains important for expressions, rules, report templates, and legacy macros. PML 2 is the object-oriented layer mainly used for GUI customization, forms, gadgets, methods, and user-defined object types."},
    @{Kind="Bullet"; Text="Variables are objects. Local variables use !name; global objects and forms commonly use !!name. Members and methods are accessed with dot notation."},
    @{Kind="Bullet"; Text="PML source is conventionally organized in PMLLIB folders. Common file roles include macros, functions, forms (.pmlfrm), object definitions (.pmlobj), and function files (.pmlfnc). After adding PML files to a running session, rebuild the PML index with pml rehash or pml rehash all."},
    @{Kind="Bullet"; Text="Methods are object-specific functions. Method signatures specify argument types and result type. The same object may expose data members, built-in methods, and user-defined methods."},
    @{Kind="Code"; Text="PML style: !result = !object.Method(arg1, arg2); !!FormName.GadgetName.Val = 'value'"},
    @{Kind="H1"; Text="Object Families in the Reference Manual"},
    @{Kind="Bullet"; Text="Built-in objects: ARRAY, BLOCK, BOOLEAN, FILE, OBJECT, REAL, STRING, DATETIME, UNIT, MEASURE."},
    @{Kind="Bullet"; Text="3D geometry objects: ARC, LINE, LINEARGRID, LOCATION, PLANE, PLANTGRID, POINTVECTOR, PROFILE, RADIAL GRID, XYPOSITION, and related geometric types."},
    @{Kind="Bullet"; Text="Plant/E3D database and session objects: BANNER, BORE, DB, DBREF, DBSESS, DIRECTION, MACRO, MDB, ORIENTATION, POSITION, POSTUNDO, PROJECT, SESSION, TEAM, UNDOABLE, USER, ALERT."},
    @{Kind="Bullet"; Text="Forms, menus, and gadgets: BAR, BUTTON, COMBOBOX, CONTAINER, FMSYS, FORM, FRAME, LINE, LIST, MENU, NUMERIC, OPTION, PARAGRAPH, RTOGGLE, SELECTOR, SLIDER, TEXT, TEXTPANE, TOGGLE, VIEW, ALPHA, AREA, PLOT, VOLUME."},
    @{Kind="Bullet"; Text="Collection and report objects: COLLECTION, COLUMN, COLUMN-FORMAT, DATE-FORMAT, EXPRESSION, REPORT, TABLE, FORMAT."},
    @{Kind="H1"; Text="Methods Common to Objects"},
    @{Kind="Bullet"; Text="Attribute('Name'): get or set a member dynamically by member name."},
    @{Kind="Bullet"; Text="Attributes(): return the member-name list for an object."},
    @{Kind="Bullet"; Text="Delete(): destroy the object or make the variable undefined, subject to object semantics."},
    @{Kind="Bullet"; Text="Comparison methods such as EQ, NEQ, and LT perform type-dependent comparison. Min and Max return the lower or higher of two comparable objects."},
    @{Kind="Bullet"; Text="String, real, array, and date/time objects expose many type-specific methods for conversion, splitting/joining, sorting, filtering, formatting, indexing, and arithmetic. Check the licensed method table for exact signatures."},
    @{Kind="H1"; Text="Arrays, Blocks, and Evaluation"},
    @{Kind="Bullet"; Text="ARRAY is the main PML collection type. It supports indexed data, iteration, sorting/filtering patterns, and bulk evaluation against BLOCK expressions."},
    @{Kind="Bullet"; Text="BLOCK stores an expression or code fragment for later evaluation, often with a loop index such as !evalIndex when mapping over arrays."},
    @{Kind="Bullet"; Text="A useful pattern is to collect DBREF objects, define a BLOCK that queries each item, then evaluate the block across the array to obtain strings, numbers, or derived data."},
    @{Kind="H1"; Text="Database and Query Objects"},
    @{Kind="Bullet"; Text="DBREF represents an element reference in the AVEVA database. DB, MDB, PROJECT, SESSION, USER, TEAM, and DBSESS expose project/session/database context."},
    @{Kind="Bullet"; Text="PML code commonly combines database references with pseudo-attributes for navigation, reports, collections, and rule evaluation."},
    @{Kind="Bullet"; Text="COLLECTION and REPORT objects support extracting sets of database elements and formatting or exporting their attributes."},
    @{Kind="Bullet"; Text="PML functions may be used as data-source hooks in other AVEVA products where the function returns a supported value and accepts no arguments."},
    @{Kind="H1"; Text="Forms and Gadgets"},
    @{Kind="Bullet"; Text="A form is a global PML object with predefined members, built-in methods, user-defined members, form variables, and gadget members. Gadgets are member objects of a form."},
    @{Kind="Bullet"; Text="Gadgets hold display state and value state. Typical members include Active, Val, Callback, geometry/anchor/docking fields, and type-specific properties."},
    @{Kind="Bullet"; Text="Callbacks are text strings containing PML expressions, function calls, method calls, commands, or form show/hide actions. They are executed when the user interacts with the form or when form lifecycle events occur."},
    @{Kind="Code"; Text="Minimal form pattern: setup form !!Hello; paragraph .Message text 'Hello'; button .Bye 'Close' OK; exit"},
    @{Kind="H1"; Text="PML 1 Expression Package"},
    @{Kind="Bullet"; Text="PML 1 expressions are still used in rules, report templates, catalogue parameterisation, and attribute-driven evaluation."},
    @{Kind="Bullet"; Text="Expression categories include numeric, logical, text, position, direction, orientation, and attribute expressions. The reference covers units, undefined/unset values, late variable evaluation, precision of comparisons, and querying expressions."},
    @{Kind="Bullet"; Text="When expressions are evaluated inside .NET through DbExpression, the caller should specify units for ambiguous numeric literals and convert raw results back to user-facing units when displaying output."},
    @{Kind="H1"; Text="PML.NET Interoperability"},
    @{Kind="Bullet"; Text="PML.NET lets PML instantiate and call callable .NET classes through proxy objects. PML imports the assembly/object exposure, sets a namespace, creates an object, and calls visible methods."},
    @{Kind="Bullet"; Text="This bridge is useful for UI widgets, file dialogs, grid controls, database helper code, and local adapters. It is not a general license to call arbitrary PML back from .NET; events/callbacks should be used for that direction."},
    @{Kind="Code"; Text="Example pattern: import 'PMLFileBrowser'; using namespace 'Aveva.Core.Presentation'; !browser = object PMLFileBrowser('OPEN')"},
    @{Kind="H1"; Text="PML File Browser Specification"},
    @{Kind="Bullet"; Text="Object: PMLFileBrowser. Constructor mode: OPEN or SAVE; if omitted, behavior defaults to open-mode in public examples."},
    @{Kind="Bullet"; Text="show(directory, seed, title, exists, filter, index): displays a Windows file-browser dialog. directory is the starting folder; seed is default filename; title is the window title; exists controls file-existence checking; filter uses description|pattern pairs; index selects active filter."},
    @{Kind="Bullet"; Text="file(): returns the selected file path after the dialog completes."},
    @{Kind="H1"; Text="Implementation Guidance"},
    @{Kind="Bullet"; Text="Keep PML object/form names short, stable, and unique. Avoid name collisions between forms, global objects, gadgets, and object types."},
    @{Kind="Bullet"; Text="Prefer PML 2 object methods for new GUI customization; use PML 1 expressions where the AVEVA product expects expression syntax for rules/reports."},
    @{Kind="Bullet"; Text="After deployment, verify PMLLIB path, index state, product module, database access, and user role before debugging the code itself."},
    @{Kind="Bullet"; Text="For external PLM integration, use PML as a host-side automation layer and place network/API integration in a .NET add-in or local service when stronger typing, authentication, or retry behavior is needed."},
    @{Kind="H1"; Text="Sources Used"},
    @{Kind="Bullet"; Text="AVEVA E3D Software Customisation Reference Manual public introduction: https://help.aveva.com/AVEVA_Everything3D/1.1/SCRM/SCRM2.html"},
    @{Kind="Bullet"; Text="AVEVA PML functions/methods documentation page/snippet: https://docs.aveva.com/bundle/engineering/page/1026679.html"},
    @{Kind="Bullet"; Text="AVEVA PML 1 and PML 2 expressions documentation page/snippet: https://docs.aveva.com/bundle/engineering/page/1026760.html"},
    @{Kind="Bullet"; Text="AVEVA PML rehash documentation page/snippet: https://docs.aveva.com/bundle/administration/page/971298.html"},
    @{Kind="Bullet"; Text="Public mirror of AVEVA Software Customisation Reference Manual, used only for paraphrased object-family extraction: https://pdfcoffee.com/aveva-software-customisation-reference-manual-pdf-free.html"}
)

$dsWeb = @(
    @{Kind="H1"; Text="Source Access Status"},
    @{Kind="Normal"; Text="The official 3DEXPERIENCE Web Services API pages require a 3DEXPERIENCE ID and/or Support access in this environment. Dassault's public developer guide confirms the official portal and access requirement. Endpoint-level detail below is reconstructed from the public Dassault Systèmes WS3DX .NET SDK repository and should be validated against the tenant's official OpenAPI files before implementation."},
    @{Kind="H1"; Text="Platform API Architecture"},
    @{Kind="Bullet"; Text="Integration is based on RESTful APIs for synchronous resource operations and JMS/event-driven messaging for asynchronous integration."},
    @{Kind="Bullet"; Text="REST resources follow standard HTTP methods such as GET, POST, PUT/PATCH-style update patterns, and DELETE. Payloads are generally JSON; STEP AP242 is identified by Dassault as a supported interchange form for relevant data exchange."},
    @{Kind="Bullet"; Text="Authentication and validation are integral. Public Dassault material mentions Basic Auth and OAuth; the SDK encapsulates Passport authentication, security context, and CSRF handling in lower layers."},
    @{Kind="Bullet"; Text="Enterprise Integration Framework patterns: event/messaging, punctual REST API operations, and asynchronous export/import operations."},
    @{Kind="Code"; Text="Common route shape from SDK: {3DSpace}/resources/v1/modeler/{family}/..."},
    @{Kind="H1"; Text="Public SDK Coverage"},
    @{Kind="Bullet"; Text="ws3dx.dseng: Engineering Web Services, SDK versions through 2025x GA 1.4.0."},
    @{Kind="Bullet"; Text="ws3dx.dsmfg: Manufacturing Item Web Services, SDK versions through 2025x GA 1.16.0."},
    @{Kind="Bullet"; Text="ws3dx.project/task: Task REST Services, SDK versions through 2025x GA 1.1.0."},
    @{Kind="Bullet"; Text="ws3dx.dsxcad: CAD Collaboration Web Services, SDK versions through 2025x GA 1.5.0."},
    @{Kind="Bullet"; Text="ws3dx.dsprcs: Manufacturing Process Web Services, SDK versions through 2025x GA 1.15.0."},
    @{Kind="Bullet"; Text="ws3dx.dsdo: Derived Outputs Web Services, SDK version 2025x GA 1.2.1."},
    @{Kind="Bullet"; Text="ws3dx.dsiss: Issue Web Services, SDK version 2025x GA 1.1.1."},
    @{Kind="H1"; Text="Engineering Web Services: dseng"},
    @{Kind="Bullet"; Text="Primary resources: Engineering Item, Engineering Instance, Engineering Representation Instance, alternates, make-from connections, configured instances, geolocation, change control, enterprise item number, deformable/deformed state, and effectivity."},
    @{Kind="Bullet"; Text="Typical operations: search, locate, get by ID, create engineering item, create instances, create representation instances, expand structure, bulk fetch/update, update item/instance, delete item/instance, attach/detach change control, set/unset evolution effectivity, set/unset variant effectivity, set/update/unset configured instance, and manage make-from connections."},
    @{Kind="Code"; Text="Base route: /resources/v1/modeler/dseng/..."},
    @{Kind="H1"; Text="Manufacturing Item Web Services: dsmfg"},
    @{Kind="Bullet"; Text="Primary resources: Manufacturing Item, Manufacturing Item Instance, alternate, substitute, origin, manufacturing responsibility, scope engineering item, resulting engineering item, partial scope engineering item, implemented engineering occurrence, assignment filter, dependency, configuration, change control, assigned requirements, and requirement specifications."},
    @{Kind="Bullet"; Text="Typical operations: search/get/create/update/delete manufacturing items, expand MBOM structures, bulk fetch, locate, create or replace instances, attach/detach scope/resulting engineering items, attach/detach implemented occurrences, connect requirements/specifications, manage alternates/substitutes, reconnect, and retrieve realized changes."},
    @{Kind="Code"; Text="Base route: /resources/v1/modeler/dsmfg/..."},
    @{Kind="H1"; Text="Manufacturing Process Web Services: dsprcs"},
    @{Kind="Bullet"; Text="Primary resources: Manufacturing Process, Manufacturing Process Instance, Manufacturing Operation, Manufacturing Operation Instance, instruction, item specification, time constraint, primary/secondary capable resource, pre-assigned work center, asset context, checklist, data-collect plan, data-collect row, resource-parameter plan, alert, and sign-off."},
    @{Kind="Bullet"; Text="Typical operations: get/search/create/update/delete processes and operations, expand process structures, attach/detach configuration and change control, set/unset evolution and variant effectivity, manage operation/process instances, manage resource assignments and time constraints, retrieve item specifications, data-collection rows, sign-off status, and realized changes."},
    @{Kind="Code"; Text="Base route: /resources/v1/modeler/dsprcs/..."},
    @{Kind="H1"; Text="CAD Collaboration Web Services: dsxcad"},
    @{Kind="Bullet"; Text="Primary resources: xCAD Part, Product, Drawing, Representation, Shape3D, Template, Family Representation, authoring files, visualization files, xCAD attributes, change control, enterprise item number, and check-in/download tickets."},
    @{Kind="Bullet"; Text="Typical operations: search/locate/get CAD objects, create or update supported resources, delete parts/drawings/representations, get authoring-file download tickets, get authoring-file check-in tickets, get visualization-file download tickets, attach/delete change control, and get/modify enterprise item numbers and xCAD attributes."},
    @{Kind="Code"; Text="Base route: /resources/v1/modeler/dsxcad/..."},
    @{Kind="H1"; Text="Derived Outputs Web Services: dsdo"},
    @{Kind="Bullet"; Text="Primary resources: Derived Output, Derived Output File, Derived Output Rule, Derived Output Job, check-in ticket, and download ticket."},
    @{Kind="Bullet"; Text="Typical operations: get derived output detail/complete masks, create derived outputs, locate derived outputs, update derived-output files, remove files, retrieve download tickets, retrieve rules, and create derived-output jobs."},
    @{Kind="Code"; Text="Base route: /resources/v1/modeler/dsdo/..."},
    @{Kind="H1"; Text="Issue Web Services: dsiss"},
    @{Kind="Bullet"; Text="Primary resource: Issue, with typed URI identifiers and issue comments."},
    @{Kind="Bullet"; Text="Typical operations exposed by the SDK: search, paged search, create, get detail, update, approve, and reject."},
    @{Kind="Code"; Text="Base route: /resources/v1/modeler/dsiss/..."},
    @{Kind="H1"; Text="Task REST Services: project/task"},
    @{Kind="Bullet"; Text="Primary resources: Task, assignee, deliverable, reference, scope, calendar, owner/originator, route, predecessor, and related data/elements."},
    @{Kind="Bullet"; Text="Typical operations: get user tasks, get tasks by payload/scope, create task, update task(s), delete task, get/add assignees, get/add deliverables, get/add references, and get/add scopes."},
    @{Kind="Code"; Text="Base route: /resources/v1/modeler/... with task-specific paths in ws3dx.project.task"},
    @{Kind="H1"; Text="Cross-Cutting Data Contract Patterns"},
    @{Kind="Bullet"; Text="Masks control which attributes a service returns. The SDK models masks as C# interfaces decorated with mask metadata and checks supported mask constraints at runtime."},
    @{Kind="Bullet"; Text="Enterprise Attributes extend standard attributes through customer key-value definitions. The SDK models these as IDictionary<string, object> derived interfaces for flexible serialization."},
    @{Kind="Bullet"; Text="Customer Attributes are similar extension attributes for private/dedicated cloud or on-premise deployments; they are not public-cloud generic attributes."},
    @{Kind="Bullet"; Text="Many service responses contain totalItems, member, and nlsLabel fields. The SDK usually unwraps the member array into IList or IEnumerable-style return values for application code."},
    @{Kind="Bullet"; Text="Known SDK limitations include incomplete $fields and $include support, missing implementation classes in some areas, and document upload/download functions still needing work in some families."},
    @{Kind="H1"; Text="Implementation Guidance"},
    @{Kind="Bullet"; Text="Always obtain the tenant-specific service URL, 3DSpace URL, security context, tenant/platform version, and official OpenAPI JSON before building production integrations."},
    @{Kind="Bullet"; Text="Use the SDK family names as a discovery map, not a warranty. The repository states it is open source collaboration material and not an official supported Dassault product."},
    @{Kind="Bullet"; Text="For AVEVA-to-3DEXPERIENCE integration, separate authoring-side local adapters from platform-side REST calls. Use IDs, typed URI identifiers, change-control links, derived outputs, and issue/task objects as the exchange contracts."},
    @{Kind="H1"; Text="Sources Used"},
    @{Kind="Bullet"; Text="Dassault Systèmes Developer Guides page: https://www.3ds.com/support/documentation/developer-guides"},
    @{Kind="Bullet"; Text="Dassault Systèmes 3DEXPERIENCE Platform Openness page: https://www.3ds.com/3dexperience/openness"},
    @{Kind="Bullet"; Text="Dassault Systèmes WS3DX .NET SDK repository: https://github.com/3ds-cpe-emed/ws3dx-dotnet"},
    @{Kind="Bullet"; Text="3DEXPERIENCE media-hosted Web Services documentation links were tested and redirected to 3DEXPERIENCE ID login in this environment."}
)

$dsCaa = @(
    @{Kind="H1"; Text="Source Access Status"},
    @{Kind="Normal"; Text="Dassault's public developer guide page confirms that CAA Encyclopedia and V5/V6 developer toolkits contain how-to articles, samples, and interface/class/method reference material, but the download/access path requires Support access and CAA media. This extraction therefore combines official public access notes with public CAA V5 Encyclopedia mirrors for the automation/site-map structure."},
    @{Kind="H1"; Text="CAA Architecture Map"},
    @{Kind="Bullet"; Text="CAA is organized by brand, solution, modeler, and framework. A solution is made up of modelers; a modeler represents the API surface of one or more products; modelers are made up of frameworks that contain APIs."},
    @{Kind="Bullet"; Text="Core foundations called out in the CAA site map: 3D PLM Enterprise Architecture, 3D PLM PPR Hub Open Gateway, and RADE."},
    @{Kind="Bullet"; Text="3D PLM Enterprise Architecture isolates application APIs from operating-system and system-software dependencies."},
    @{Kind="Bullet"; Text="3D PLM PPR Hub Open Gateway provides Process, Product, and Resource modelers and gateway mechanisms for exchanging data with CAD systems and standard formats."},
    @{Kind="Bullet"; Text="RADE supplies the development environment and tools to design, implement, build, check, test, and manage CAA application sources."},
    @{Kind="Bullet"; Text="CAA documentation has two major tracks: C++/Java APIs and CAA Automation. The automation track is the COM/IDL scripting interface used from VB/VBA and other COM-capable languages."},
    @{Kind="H1"; Text="CAA Automation Foundation"},
    @{Kind="Bullet"; Text="The Infrastructure Automation home page organizes articles on object architecture, object diagrams, collections, SafeArrayVariant handling, sub/function procedures, numerical values and units, invoking CATIA from scripting languages, VB/VBA portability, and settings controllers."},
    @{Kind="Bullet"; Text="Infrastructure use cases include macro recording/replay, adding macros to toolbars, file/folder access, creating a document, opening a document, closing a document, saving, and Save As."},
    @{Kind="Bullet"; Text="Quick references include Infrastructure Automation objects, System Automation objects, and setting-controller objects/property pages/types."},
    @{Kind="Code"; Text="Typical automation root: CATIA.Application -> Documents -> PartDocument/ProductDocument/DrawingDocument"},
    @{Kind="H1"; Text="Common Automation Object Areas"},
    @{Kind="Bullet"; Text="Application/session: Application, Windows, Documents, ActiveDocument, Selection, SystemService, FileSystem, and setting controllers."},
    @{Kind="Bullet"; Text="Part design: PartDocument, Part, Bodies, Body, HybridBodies, HybridShapeFactory, ShapeFactory, Parameters, Relations, Sketches, Constraints, and update methods."},
    @{Kind="Bullet"; Text="Assembly: ProductDocument, Product, Products collection, components, constraints, publications, positioning, replacement, and assembly holes."},
    @{Kind="Bullet"; Text="Drafting: DrawingDocument, DrawingSheets, DrawingSheet, DrawingViews, DrawingView, tables, dimensions, generative views, details, print areas, and sheet update operations."},
    @{Kind="Bullet"; Text="DMU/measurement: workbenches such as SPAWorkbench, measurable objects, clash/space analysis families where licensed, and visualization/navigation objects."},
    @{Kind="H1"; Text="C++/Java CAA API Layer"},
    @{Kind="Bullet"; Text="The native CAA layer is compiled and framework-based. It is used for deeper product extensions where Automation macros are not sufficient: custom commands, features, dialogs, modelers, data exchange, lifecycle integration, and high-performance domain logic."},
    @{Kind="Bullet"; Text="APIs are grouped into frameworks within modelers. Implementations generally follow Dassault interface/component conventions, with framework dependencies declared in the RADE/build environment."},
    @{Kind="Bullet"; Text="Authorized APIs vary by release and installed products. The CAA What's New page lists new and changed authorized APIs by release, including modelers and solutions added across V5 releases."},
    @{Kind="H1"; Text="Mechanical, Assembly, and Drafting Guide Surface"},
    @{Kind="Bullet"; Text="Mechanical Design automation includes product modelers such as Part Design and related mechanical modeler references."},
    @{Kind="Bullet"; Text="Assembly automation exposes assembly feature objects and use cases such as creating constraints on published elements and creating/modifying assembly holes."},
    @{Kind="Bullet"; Text="Drafting automation exposes drawing objects and use cases including creating front views, updating drawing sheets, creating sheets/detail sheets, creating drawing tables, checking dimensions, duplicating views, instantiating details, and setting print areas."},
    @{Kind="H1"; Text="Build and Deployment Specification"},
    @{Kind="Bullet"; Text="Official CAA installation guidance is distributed by Dassault and covers installing CATIA V5 and CAA APIs plus running/debugging sample use cases. The public page advertises a downloadable PDF but detailed download requires the normal 3DS documentation flow."},
    @{Kind="Bullet"; Text="CAA developer guides are automatically installed when CAA media is installed. Dassault's public developer page points users to the software download platform for CAA media and states Support access is required for developer guide access."},
    @{Kind="Bullet"; Text="Automation macros are lighter-weight and can be distributed as CATScript/VBA/CATVBA or external COM automation scripts. CAA compiled extensions require the matching release/toolchain/runtime environment."},
    @{Kind="H1"; Text="Implementation Guidance"},
    @{Kind="Bullet"; Text="Use Automation for repeatable document/model operations, data extraction, drafting automation, and office-style macros. Use native CAA C++/Java when you need compiled commands, custom model features, deeper UI integration, or higher-performance product extensions."},
    @{Kind="Bullet"; Text="Treat every object reference as release/product/license dependent. The installed CAA Encyclopedia for the target CATIA/3DEXPERIENCE release is the authority for exact interfaces, methods, and supported frameworks."},
    @{Kind="Bullet"; Text="For integration with PLM or AVEVA workflows, keep CATIA/CAA logic on the authoring side and expose a controlled adapter/API boundary to enterprise services."},
    @{Kind="H1"; Text="Sources Used"},
    @{Kind="Bullet"; Text="Dassault Systèmes Developer Guides page: https://www.3ds.com/support/documentation/developer-guides"},
    @{Kind="Bullet"; Text="Dassault Systèmes Installation of CATIA V5 and CAA API page: https://www.3ds.com/support/documentation/resource-library/installation-catia-v5-and-caa-api"},
    @{Kind="Bullet"; Text="Public CAA V5 Automation Site Map mirror: https://www.maruf.ca/files/caadoc/CAAScdBase/CAAScdAutomationSiteMap.htm"},
    @{Kind="Bullet"; Text="Public Infrastructure Automation home page mirror: https://www.maruf.ca/files/caadoc/CAAScdBase/CAAInfScriptHome.htm"},
    @{Kind="Bullet"; Text="Public Assembly Automation home page mirror: https://www.maruf.ca/files/caadoc/CAAScdBase/CAAAsmScriptHome.htm"},
    @{Kind="Bullet"; Text="Public Drafting Automation home page mirror: https://www.maruf.ca/files/caadoc/CAAScdBase/CAADriScriptHome.htm"},
    @{Kind="Bullet"; Text="Public CAA What's New mirror: https://www.maruf.ca/files/caadoc/CAACenQuickRefs/CAACenWhatsNew.htm"}
)

$word = New-WordApp
$created = @()
try {
    $created += Build-Doc $word "AVEVA Everything3D .NET Customisation Guide" $dateNote $avevaNet "AVEVA_Everything3D_NET_Customisation_Key_API_Specifications_EN"
    $created += Build-Doc $word "AVEVA PML Reference Manual" $dateNote $avevaPml "AVEVA_PML_Reference_Manual_Key_API_Specifications_EN"
    $created += Build-Doc $word "Dassault 3DEXPERIENCE Platform Web Services API Reference" $dateNote $dsWeb "Dassault_3DEXPERIENCE_Web_Services_Key_API_Specifications_EN"
    $created += Build-Doc $word "Dassault CAA Guide" $dateNote $dsCaa "Dassault_CAA_Guide_Key_API_Specifications_EN"
} finally {
    if ($word.Documents.Count -eq 0) {
        $word.Quit()
    }
}

$created | Format-Table -AutoSize
