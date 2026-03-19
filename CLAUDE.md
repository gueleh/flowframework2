# FlowFramework 2 – Claude Code Instructions

## Overview

FlowFramework 2 (FF2) is a VBA framework for building Excel applications (.xlsb). Code is exported from Excel as `.bas` (modules), `.cls` (classes) and `.frm` (forms) files. Claude Code works only with these exported code files — the export/reimport cycle into Excel is handled by the user.

**Never modify files marked `' CORE, do not change`** — these are framework internals. Only modify files belonging to the application layer (`a_` prefix) or the customizable app-framework layer (`af_` prefix, where marked with `>>>>>>> Your code/cases here`). The only exception: the user specifies explicitly that the framework itself is supposed to be changed.

## Architecture & Layers

The framework has three layers, identified by file name prefixes:

| Prefix | Layer | Purpose | Editable? |
|--------|-------|---------|-----------|
| `f_` | Framework Core | Core framework functionality | NO (CORE) |
| `af_` | App-Framework | Customizable framework parts (error handling, globals, modes) | Only within `>>>>>>>` and `<<<<<<<` markers |
| `a_` | Application | Application-specific code (your code lives here) | YES |
| `DEV_f_` | Dev Framework | Development/testing tools, removed on deployment | NO (CORE-DEV) |
| `DEV_a_` | Dev Application | Application dev/test code | YES |

### Module Type Prefixes (after layer prefix)

| Suffix | Type | Example |
|--------|------|---------|
| `pM_` | Private Module (`Option Private Module`) | `a_pM_EntryLevel.bas` |
| `M_` | Public Module | `f_M_TemplRenderer_Types.bas` |
| `C_` | Class Module | `f_C_CallParams.cls` |
| `I_` | Interface Class | `f_I_DataRecord.cls` |
| `wks_` | Worksheet code-behind | `a_wks_Main.cls` |
| `wkb_` | Workbook code-behind | `a_wkb_Main.cls` |

### Key Files for Application Development

- **`a_pM_EntryLevel.bas`** — Add entry-level Subs here (user-facing actions)
- **`a_pM_Globals.bas`** — Application-specific global variables
- **`a_C_AppSettings.cls`** — Application settings class
- **`a_pM_OnChangeSubsFor_f_C_Wks.bas`** — Worksheet change event handlers
- **`af_pM_ErrorHandling.bas`** — Add app-specific error enum cases and descriptions
- **`af_pM_Globals.bas`** — Add custom processing modes and settings sheets
- **`af_C_AppModes.cls`** — Customize maintenance/dev sheet visibility lists

## Variable Naming Convention (Hungarian Notation)

Every variable name starts with a type prefix:

| Prefix | Type | Example |
|--------|------|---------|
| `s` | String | `sName`, `sPassword` |
| `l` | Long | `lRow`, `lColumn`, `lIndex` |
| `b` | Boolean | `bFound`, `bSuccess` |
| `o` | Object (generic) | `oWks` (Worksheet), `oWkb` (Workbook), `oRng` (Range) |
| `oC` | Object (custom class) | `oC_Me` (f_C_CallParams), `oC_Error` (f_C_Error) |
| `oCol` | Collection | `oCol_Errors` |
| `oDict` | Dictionary | `oDict_ColumnsByName` |
| `v` | Variant | `vItem`, `vValue` |
| `va` | Variant Array | `vaData()`, `vaValues()` |
| `sa` | String Array | `saTokens()`, `saParts()` |
| `la` | Long Array | `laIndexes()` |
| `e` | Enum value | `eVisibility`, `eProcessingMode` |
| `u` | User-Defined Type | `uBlockSpec`, `uCellSpec` |
| `dte` | Date | `dteVersionDate` |

### Scope Modifiers (inserted between type prefix and name)

| Modifier | Scope | Example |
|----------|-------|---------|
| `_m_` | Private module/class-level variable | `s_m_ComponentName`, `l_m_Index` |
| `_arg_` | Parameter (argument) | `s_arg_Name`, `l_arg_Row` |
| `_p_` or `_f_p_` | Public framework-level | `s_f_p_ERROR`, `oC_f_p_FrameworkSettings` |

### Naming for Constants

- Private module constant: `s_m_COMPONENT_NAME`, `s_m_KEY_HEADER` (UPPER_CASE after scope prefix)
- Public constant: `s_f_p_SPLIT_SEED_SEPARATOR`, `s_f_p_ERROR`

### Property Naming Convention

Properties encode their access level in the name:

```vba
' Read-Write property (prop_rw_)
Public Property Get s_prop_rw_ComponentName() As String
Public Property Let s_prop_rw_ComponentName(ByVal sNewValue As String)

' Read-Only property (prop_r_)
Public Property Get oC_prop_r_Error() As f_C_Error
```

### Function/Sub Naming Convention

Function and Sub names encode the return type and scope:

```vba
' Boolean-returning framework-public function
Public Function b_f_p_GetWorkbook() As Boolean

' String-returning framework-public function
Public Function s_f_p_HandledErrorDescription() As String

' Variant-returning function
Public Function v_f_p_ValueFromWorkbookName() As Variant

' Object-returning function (prefix = object type abbreviation)
Public Function oWks_f_p_WorksheetFromCodeNameString() As Worksheet
Public Function oRng_f_p_RangeFromWorkbookName() As Range

' Sub (no return type prefix)
Public Sub f_p_StartProcessing()

' Private method (m prefix instead of p)
Private Sub mLogError()
Private Function b_m_WksIsNotSet() As Boolean
```

## Procedure Templates — THE CORE PATTERNS

Every non-trivial procedure MUST follow one of these templates. The templates exist in two forms: **standard** (with detailed comments) and **compact**. Use compact for production code.

### Entry-Level Sub (Compact Template)

Entry-level Subs are the top-level procedures triggered by user actions (buttons, menus). They are the ONLY place where `b_prop_rw_SilentError = False` should be used. These always are supposed to be placed in a module such as a_pM_EntryLevel.bas - i.e. a private Module specified as entry level. Thus a button cannot call such a sub, for this purpose a trival user facing sub is used in a public module such as a_M_UserInterface.bas, only containing one line of code which calls the entry level sub. The block after "Finally:" must be executed, regardless of what happens during execution - the "Try:", "Finally:", "HandleError:" ,"Catch:" logic simulates Try-Catch-Finally of other, higher programming languages - with "HandleError:" as additional block, which is required in case of error specific stuff to be done before "Finally:", but after "Catch:".

```vba
Public Sub a_p_MyEntryLevelSub()
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME
'>>>>>>> Your custom settings here
   f_p_StartProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
   With oC_Me
      .s_prop_rw_ProcedureName = "a_p_MyEntryLevelSub"
      .b_prop_rw_SilentError = False
      .s_prop_rw_ErrorMessage = "Descriptive error message for the user."
      .SetCallArgs "No args"
   End With
'Fixed, don't change
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
Try: On Error GoTo Catch


'>>>>>>> Your code here

   ' Call lower-level functions like this:
      If Not _
   b_a_p_MyLowerLevelFunction() _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)

'End of your code <<<<<<<


'Fixed, don't change
Finally: On Error Resume Next


'>>>>>>> Your code here
   ' Cleanup code (always executed)

'End of your code <<<<<<<


'>>>>>>> Your custom settings here
   f_p_EndProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
'Fixed, don't change
   Exit Sub
HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_EntryLevel


'>>>>>>> Your code here
   ' Error-specific cleanup

'End of your code <<<<<<<


'Fixed, don't change
   Resume Finally
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: GoTo HandleError
   End If
End Sub
```

### Lower-Level Function (Compact Template)

Lower-level functions return `Boolean` — `True` means success, `False` (default) means failure. Return values are passed via `ByRef` parameters, assigned in the `Finally` block. The block after "Finally:" must be executed, regardless of what happens during execution - the "Try:", "Finally:", "HandleError:" ,"Catch:" logic simulates Try-Catch-Finally of other, higher programming languages - with "HandleError:" as additional block, which is required in case of error specific stuff to be done before "Finally:", but after "Catch:". These functions must be called either by a Entry Level sub as shown above or by another Lower-Level Function using this structure - otherwise the framework logic would not works, especially regarding error handling and logging.

```vba
Public Function b_a_p_MyLowerLevelFunction(ByRef sOutput As String, ByVal sInput As String) As Boolean
'Fixed, don't change
   Dim oC_Me As New f_C_CallParams: oC_Me.s_prop_rw_ComponentName = s_m_COMPONENT_NAME: If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterUnitTest oC_Me
'>>>>>>> Your custom settings here
   With oC_Me
      .s_prop_rw_ProcedureName = "b_a_p_MyLowerLevelFunction"
      .b_prop_rw_SilentError = True
      .s_prop_rw_ErrorMessage = "Descriptive error message."
      .SetCallArgs "sInput:=" & sInput
   End With
'Fixed, don't change
Try: On Error GoTo Catch


'>>>>>>> Your code here
   Dim sTempOutput As String
   sTempOutput = "processed: " & sInput

'End of your code <<<<<<<


'Fixed, don't change
Finally: On Error Resume Next


'>>>>>>> Your code here
   sOutput = sTempOutput   ' Assign ByRef output in Finally block

'End of your code <<<<<<<


'>>>>>>> Your custom settings here
   If oC_Me.oC_prop_r_Error Is Nothing Then b_a_p_MyLowerLevelFunction = True
'Fixed, don't change
   Exit Function
HandleError: af_pM_ErrorHandling.af_p_Hook_ErrorHandling_LowerLevel


'>>>>>>> Your code here

'End of your code <<<<<<<


'Fixed, don't change
   Resume Finally
Catch:
   If oC_Me.oC_prop_r_Error Is Nothing Then f_p_RegisterError oC_Me, Err.Number, Err.Description
   If oC_f_p_FrameworkSettings.b_prop_rw_ThisIsATestRun Then f_p_RegisterExecutionError oC_Me
   If oC_f_p_FrameworkSettings.b_prop_r_DebugModeIsOn And Not oC_Me.b_prop_rw_ResumedOnce Then
      oC_Me.b_prop_rw_ResumedOnce = True: Stop: Resume
   Else
      f_p_HandleError oC_Me: GoTo HandleError
   End If
End Function
```

## Critical Pattern Details

### Calling Lower-Level Functions (Error Propagation)

Always use this indentation pattern when calling lower-level functions:

```vba
      If Not _
   b_a_p_SomeLowerLevelFunction(sResult, sParam1) _
      Then Err.Raise _
         e_f_p_HandledError_ExecutionOfLowerLevelFunction, , _
         s_f_p_HandledErrorDescription(e_f_p_HandledError_ExecutionOfLowerLevelFunction)
```

The specific indentation (function name left-aligned with surrounding code, `If Not` and `Then Err.Raise` indented further) is deliberate — it makes lower-level calls visually distinct.

### ByRef Return Values

- The Boolean return value indicates success/failure
- Actual output values are passed via `ByRef` parameters
- Always assign `ByRef` outputs in the `Finally` block (ensures they are set even on error)
- Use temporary variables in the `Try` block, assign to `ByRef` params in `Finally`

```vba
Public Function b_a_p_GetData(ByRef vaResult() As Variant, ByVal lSheetIndex As Long) As Boolean
   ' ...
Try: On Error GoTo Catch
   Dim vaTemp() As Variant
   ' ... populate vaTemp ...

Finally: On Error Resume Next
   vaResult = vaTemp   ' Assign ByRef in Finally
   If oC_Me.oC_prop_r_Error Is Nothing Then b_a_p_GetData = True
   Exit Function
   ' ...
End Function
```

### Error Handling Flow

```
Try → (error occurs) → Catch → f_p_RegisterError → f_p_HandleError → HandleError hook → Resume Finally → cleanup → Exit
                                                                                                                     ↑
Try → (no error) → Finally → cleanup → Boolean = True → Exit ───────────────────────────────────────────────────────┘
```

**Key differences between Entry-Level and Lower-Level:**
- Entry-Level: `b_prop_rw_SilentError = False` (shows MsgBox), uses `af_p_Hook_ErrorHandling_EntryLevel`, calls `f_p_StartProcessing`/`f_p_EndProcessing`
- Lower-Level: `b_prop_rw_SilentError = True` (silent), uses `af_p_Hook_ErrorHandling_LowerLevel`, NO `f_p_StartProcessing`/`f_p_EndProcessing`

### Processing Modes

Entry-level Subs must call `f_p_StartProcessing` at the beginning and `f_p_EndProcessing` in the `Finally` block:

```vba
' Default: disable screen updating and auto-calculation
f_p_StartProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn
' ... in Finally:
f_p_EndProcessing e_f_p_ProcessingMode_AutoCalcOffOnSceenUpdatingOffOn

' Lightweight: only initialize globals
f_p_StartProcessing   ' no args

' App-specific mode (define in af_pM_Globals.bas)
f_p_StartProcessing e_f_p_ProcessingMode_AppSpecific, e_af_p_ProcessingModeYourMode
```

### The s_m_COMPONENT_NAME Constant

Every module and class MUST declare this constant at the top:

```vba
Private Const s_m_COMPONENT_NAME As String = "a_pM_EntryLevel"
```

It is used by `f_C_CallParams` for error tracking and logging.

### SetCallArgs

Document procedure arguments for error logging:

```vba
.SetCallArgs "No args"
.SetCallArgs "sName:=" & sName, "lRow:=" & lRow
```

## Module File Header Template

```vba
Attribute VB_Name = "a_pM_MyModule"
' Belongs to APP - will not be updated when updating the framework
'============================================================================================
'   NAME:     a_pM_MyModule
'============================================================================================
'   Purpose:  <description>
'   Access:   Private
'   Type:     Module
'   Author:   <name>
'   Contact:  <email>
'   GitHubID: <id>
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.1.0    DD.MM.YYYY    <dev>    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "a_pM_MyModule"
```

### Class File Header Template

```vba
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a_C_MyClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Belongs to APP - will not be updated when updating the framework
'============================================================================================
'   NAME:     a_C_MyClass
'============================================================================================
'   Purpose:  <description>
'   Access:   Public
'   Type:     Class Module
'   Author:   <name>
'   Contact:  <email>
'   GitHubID: <id>
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   0.1.0    DD.MM.YYYY    <dev>    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "a_C_MyClass"
```

### Procedure Comment Header

```vba
' Purpose: <what it does>
' <version>    <date>    <developer>    <change description>
Public Function b_a_p_MyFunction() As Boolean
```

## Key Framework Classes (Reference)

### f_C_CallParams
Instantiated in every non-trivial procedure. Stores metadata for error handling and testing.
- Properties: `s_prop_rw_ComponentName`, `s_prop_rw_ProcedureName`, `b_prop_rw_SilentError`, `s_prop_rw_ErrorMessage`, `oC_prop_r_Error` (read-only), `b_prop_rw_ResumedOnce`
- Methods: `SetCallArgs(ParamArray)`, `SetoCError(f_C_Error)`, `sArgsAsString()`

### f_C_Settings (Global: `oC_f_p_FrameworkSettings`)
Framework settings singleton, initialized via `f_p_InitGlobals`.
- `b_prop_r_DebugModeIsOn`, `b_prop_r_DevelopmentModeIsOn`, `b_prop_r_MaintenanceModeIsOn`
- `b_prop_rw_ThisIsATestRun`

### f_C_Wks
Enhanced worksheet wrapper with data range management and header dictionaries.
- `Construct(oWks)` — initialize with a worksheet
- `SetDataRangeByAnchors(oRngTopLeft, [oRngTopRight], [bFirstRowIsHeader], [bCreateHeaderDict])`
- `l_prop_r_ColumnNumberByHeaderName(sHeaderName)` — column lookup by header
- `oRng_prop_r_Data`, `oRng_prop_r_DataWithoutHeader`
- Fires worksheet change events via `Application.Run` when `b_prop_rw_WksChangeEventIsActive = True`

### f_C_DataRecord / f_I_DataRecord
Dictionary-based data record with interface for polymorphism.
- `bGetFieldValue(sFieldName, ByRef vValue) As Boolean`
- `bSetFieldValue(sFieldName, vValue) As Boolean`

### f_C_SettingsSheet
Reads key-value settings from structured worksheets.

### f_C_RangeArrayProcessor
Utility for processing range data as arrays.

## Error Handling Customization

Add app-specific error cases in `af_pM_ErrorHandling.bas`:

```vba
Public Enum e_af_p_HandledErrors
   e_af_p_HandledError_GeneralError = 19999
   ' Add your cases:
   e_af_p_HandledError_InvalidInput
   e_af_p_HandledError_DataNotFound
End Enum
```

Then add descriptions in `s_af_p_HandledErrorDescription`:

```vba
Case e_af_p_HandledError_InvalidInput
   sDesc = "The provided input is invalid."
```

Use custom errors via:

```vba
Err.Raise e_f_p_HandledError_AppSpecificError, , _
   s_f_p_HandledErrorDescription(e_f_p_HandledError_AppSpecificError, e_af_p_HandledError_InvalidInput)
```

## Global Variables (Framework)

Defined in `f_pM_GlobalsCore.bas`:
- `oC_f_p_FrameworkSettings As f_C_Settings` — framework settings instance
- `oCol_f_p_Errors As Collection` — error collection for current execution
- `oCol_f_p_UnitTests As Collection` — unit test collection

## Available References

The project uses these COM references (from `References.json`):
- VBA, Excel 16.0, OLE Automation, Office 16.0
- **Microsoft Scripting Runtime** (`Scripting.Dictionary`)
- **Microsoft VB for Applications Extensibility 5.3** (VBIDE, for code export)
- **Windows Script Host Object Model** (IWshRuntimeLibrary, for file system)

## Named Ranges and Worksheets

See `Names.json` for all named ranges and `WorksheetNames.json` for worksheet CodeNames. Worksheets are referenced by their CodeName in VBA (e.g., `a_wks_Main`, `af_wks_ErrorLog`), which differs from the tab name visible in Excel.

Settings sheets follow a structured format documented in `SettingsSheet-*.json` files.

## Coding Rules Summary

1. **Always use `Option Explicit`** in every module
2. **Always declare `s_m_COMPONENT_NAME`** as the first constant
3. **Use `Option Private Module`** for all `pM_` modules
4. **Follow Hungarian notation** for all variable names
5. **Non-trivial procedures must use the template pattern** with `f_C_CallParams`, `Try/Finally/Catch`
6. **Entry-level Subs**: `SilentError = False`, use `f_p_StartProcessing`/`f_p_EndProcessing`, `af_p_Hook_ErrorHandling_EntryLevel`
7. **Lower-level Functions**: `SilentError = True`, return `Boolean`, use `ByRef` for outputs, `af_p_Hook_ErrorHandling_LowerLevel`
8. **Assign ByRef parameters in the `Finally` block**
9. **Use the indentation pattern** for lower-level function calls with `If Not ... Then Err.Raise`
10. **Never modify CORE files** — only work in `a_` and `af_` (within markers) files
11. **Trivial helper functions** (pure logic, no side effects) may omit the full template but should still follow naming conventions
12. **Properties use the `prop_rw_` / `prop_r_` naming pattern**
13. **Use `Scripting.Dictionary` for key-value lookups** (reference is available)
14. **Use the existing framework utilities** (`f_pM_Utilities`, `f_pM_UtilitiesRanges`, etc.) rather than reimplementing
15. **New app modules** should use the `a_pM_` or `a_C_` prefix
16. **The folders `ff2s-little-sis-DEPRECATED` and `independent-features-DEPRECATED` should be ignored**

## Clean Code & Clean Architecture Principles (adapted for VBA / FF2)

The following principles are derived from Robert C. Martin's *Clean Code* and *Clean Architecture*, adapted to the realities of Excel VBA and FlowFramework 2. They complement — and never override — the framework patterns and naming conventions defined above.

### 1. Meaningful Names

- **Intention-revealing names.** A name should tell *why* something exists, *what* it does, and *how* it is used. Hungarian prefixes already encode the type; the rest of the name must encode the *meaning*. `lRow` is acceptable for a generic loop counter, but `lFirstDataRow` or `lCustomerRow` is better when context matters.
- **Avoid disinformation.** Do not call a `Collection` variable `oDict_...` or a `Dictionary` variable `oCol_...`. The prefix must match the actual type.
- **Use pronounceable, searchable names.** Prefer `sCustomerName` over `sCNm`. Single-letter names (`l`, `s`) are acceptable only in tiny scopes (loop counters, lambda-like helpers).
- **One word per concept.** Pick one term for each abstract concept and stick with it across the entire application layer. If you call it `Fetch` in one module, don't call it `Get` or `Retrieve` in another for the same semantic operation.

### 2. Functions / Procedures

- **Small.** Every procedure should do one thing and do it well. If a procedure needs a comment to explain *what* it does (beyond the `Purpose:` header), it is probably too long or doing too much.
- **Do one thing (Single Responsibility at the procedure level).** A function that validates input should not also format output. Split these into separate lower-level functions.
- **One level of abstraction per function.** An entry-level Sub should read like a high-level description of the workflow: call `b_a_p_ValidateInput`, then `b_a_p_ProcessData`, then `b_a_p_WriteOutput`. Low-level details (cell manipulation, string parsing) belong in the called functions, not in the caller.
- **Descriptive names.** A long descriptive name is better than a short cryptic one. `b_a_p_BuildHeaderDictionaryFromRange` is better than `b_a_p_BuildDict`.
- **Few arguments.** Aim for zero to three parameters. When more are needed, consider grouping related values into a UDT (`Type`) or a class. VBA does not support passing structured literals, so a dedicated parameter class or UDT is the idiomatic alternative.
- **No side effects.** A function named `b_a_p_ValidateInput` should not silently modify a worksheet. If it must change state, the name should reflect that (e.g., `b_a_p_ValidateAndMarkInput`), or the mutation should be performed by the caller based on the function's return value.
- **Command-Query Separation.** A function should either return information (query) or change state (command), not both. In FF2 terms: lower-level functions return `Boolean` + `ByRef` outputs (query-like), while Subs perform actions (commands). Keep this distinction sharp.

### 3. Comments

- **Code should be self-explanatory.** Good names and small functions reduce the need for comments. The FF2 procedure header (`Purpose:`, version history) is mandatory; additional inline comments should explain *why*, never *what*.
- **Don't comment bad code — rewrite it.** If you need a comment to explain convoluted logic, simplify the logic instead.
- **Avoid redundant comments.** Do not write `' Increment counter` above `lCounter = lCounter + 1`.
- **Acceptable comments:** Legal/license headers, `Purpose:` headers, `TODO:` / `BACKLOG:` items, warnings of consequences (e.g., `' This clears the entire output range`), clarification of obscure VBA/COM behaviour.
- **Avoid commented-out code.** Delete it. Version control preserves history.

### 4. Formatting

- **Vertical openness between concepts.** Separate logically distinct blocks with blank lines. Related statements should stay together.
- **Vertical ordering.** Higher-level procedures should appear before the lower-level functions they call (caller above callee). Within a module, organise: constants → module-level variables → public entry-level Subs → public lower-level functions → private helpers.
- **Keep lines reasonably short.** VBA's line continuation (`_`) is unwieldy; if a statement needs many continuations, consider intermediate variables. Exception: the `If Not ... Then Err.Raise` call pattern is deliberately multi-line and should follow the prescribed indentation exactly.
- **Consistent indentation.** Use three spaces for indentation (FF2 convention). Never mix tabs and spaces.

### 5. Error Handling

FF2 already provides a robust error handling framework. The Clean Code additions:

- **Don't return error codes — use the framework pattern.** Every non-trivial function uses the `Try/Finally/HandleError/Catch` template. Never invent alternative error signalling (e.g., returning `""` or `-1` to indicate failure).
- **Write informative error messages.** `s_prop_rw_ErrorMessage` should contain enough context for a user or developer to understand what failed. Include relevant parameter values or state descriptions. Example: `"Failed to read customer data from sheet '" & sSheetName & "'."` instead of `"Error reading data."`.
- **Don't suppress errors silently.** Within the `HandleError` block, you may add error-specific cleanup, but never swallow errors by doing nothing. The framework handles propagation — trust it.
- **Define error cases at the right level.** Use `e_af_p_HandledError_...` enums for application-specific errors. Keep descriptions user-friendly in entry-level contexts and developer-friendly in lower-level contexts.

### 6. Classes and Modules — Cohesion and Coupling

- **High cohesion.** Every module or class should have a single, clearly defined responsibility. A module named `a_pM_CustomerImport` should contain only procedures related to importing customer data, not unrelated utility functions.
- **Low coupling.** Modules should depend on abstractions (interfaces like `f_I_DataRecord`), not on concrete implementations. Pass required data as parameters rather than relying on shared global state beyond the framework-provided globals.
- **Keep classes small.** A class with 20+ properties or methods probably has more than one responsibility. Consider splitting it.
- **Encapsulation.** In VBA, use `Private` member variables with `Property Get`/`Let`/`Set` accessors. Never expose module-level variables as `Public` when a property would work. Use `Option Private Module` on all private modules (enforced by the `pM_` naming convention).

### 7. The Dependency Rule (Clean Architecture)

FF2's three-layer architecture (`f_` → `af_` → `a_`) embodies the Dependency Rule:

- **Dependencies point inward.** Application code (`a_`) may call framework code (`f_`), but framework code must never reference application code. The `af_` layer is the designated boundary for customization.
- **The framework is a detail, not the centre.** Application logic should not be shaped by VBA/Excel limitations. Isolate Excel-specific operations (Range reads, worksheet access) in dedicated modules. Business logic modules should receive data as arrays, dictionaries or class instances — not as Range objects.
- **Boundaries are explicit.** The `>>>>>>>` / `<<<<<<<` markers in `af_` files are architectural boundaries. Respect them: inject application-specific behaviour only at these extension points.
- **Worksheets are I/O.** Think of worksheets as external interfaces (like a database or API). Read data from worksheets into VBA data structures at the boundary, process it in pure VBA logic, then write results back. This keeps business logic testable and independent of the worksheet layout.

### 8. The SOLID Principles in VBA

#### S — Single Responsibility Principle (SRP)
Every module and class should have exactly one reason to change. If `a_pM_CustomerImport.bas` handles both data validation and database writing, split it into `a_pM_CustomerValidation.bas` and `a_pM_CustomerPersistence.bas`.

#### O — Open/Closed Principle (OCP)
Design modules to be *open for extension, closed for modification*. FF2's `af_` layer with `>>>>>>>` markers is a direct application of OCP: the framework is closed for modification, but open for extension at defined points. In application code, prefer adding new procedures or modules over modifying existing, stable ones.

#### L — Liskov Substitution Principle (LSP)
When implementing interfaces (`f_I_DataRecord`, custom `a_I_...` interfaces), every implementation must be usable wherever the interface is expected without breaking behaviour. A `a_C_CustomerRecord` implementing `f_I_DataRecord` must honour the contract: `bGetFieldValue` returns `True` and a valid value for known fields, `False` for unknown ones.

#### I — Interface Segregation Principle (ISP)
Keep interfaces small and focused. If a class only needs to read data, don't force it to implement write methods. Create separate interfaces for distinct capabilities (e.g., `a_I_Readable`, `a_I_Writable`) rather than one monolithic interface.

#### D — Dependency Inversion Principle (DIP)
High-level modules should not depend on low-level modules; both should depend on abstractions. In VBA, this means:
- Declare variables as interface types (`Dim oC_Record As f_I_DataRecord`) rather than concrete classes where possible.
- Pass dependencies as parameters instead of creating them inside the procedure.
- Use `f_C_Wks` (the framework's worksheet abstraction) instead of directly manipulating `Worksheet` objects where possible.

### 9. Code Smells to Avoid

| Smell | VBA / FF2 Context | Remedy |
|-------|-------------------|--------|
| **Long procedure** | A function exceeding ~40 lines of business logic (excluding template boilerplate) | Extract into multiple lower-level functions |
| **Large module** | A `.bas` file with 10+ unrelated procedures | Split by responsibility into separate modules |
| **Duplicate code** | Same Range-reading logic in multiple modules | Extract to a shared `a_pM_...` utility module or parameterised function |
| **Magic numbers/strings** | Hard-coded row numbers, column letters, sheet names | Use named ranges (`Names.json`), constants, or `f_C_SettingsSheet` |
| **Global mutable state** | Application-level variables modified from everywhere | Minimise globals; pass data explicitly. Use `a_pM_Globals.bas` sparingly |
| **Feature envy** | A procedure in module A that mostly accesses data from class B | Move the procedure to class B or pass the needed data as parameters |
| **Dead code** | Commented-out blocks, unreachable procedures | Delete. Git keeps the history |
| **Inconsistent abstraction level** | A procedure that mixes `Range.Value` reads with business logic | Separate I/O (worksheet access) from computation |

### 10. The Boy Scout Rule

> *Leave the code cleaner than you found it.*

When modifying a module, improve small things in passing: fix a misleading variable name, remove a dead variable, clarify a comment. Keep such cleanups minimal and confined to the module you are already changing — do not refactor unrelated code without explicit instruction.

### 11. Testing (in FF2 Context)

- **Write testable functions.** Functions that take input parameters and return results via `ByRef` / `Boolean` are inherently testable. Avoid procedures that can only be tested by inspecting worksheet state.
- **Use the framework's test infrastructure.** `f_p_RegisterUnitTest` and `b_prop_rw_ThisIsATestRun` exist for a reason. New lower-level functions should be callable during test runs.
- **Test edge cases.** Empty ranges, missing sheets, zero-length strings, uninitialised variants — VBA has many ways to fail silently. Validate inputs at the boundary.

### 12. Successive Refinement

First make it work, then make it clean:

1. **Make it work** — Write the procedure following the FF2 template. Get the logic correct.
2. **Make it right** — Refactor: extract sub-functions, rename variables for clarity, remove duplication.
3. **Make it small** — Each procedure should be as short as possible while remaining readable. If the `Try` block of a lower-level function exceeds ~30 lines of business logic, look for extraction opportunities.
