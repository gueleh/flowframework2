# Developer Manual for Flow Framework 2
Most of the documentation you can find directly in the code.

## Naming Conventions
### General Syntax
* prefixes are in lower case
* variable names are in CamelCase
* constants are in `UPPER_CASE_SEPARATED_BY_UNDERSCORE`
* always begin with element membership indicator for `Codename` properties of VBA objects, i.e. workbooks, worksheets, modules, class modules (see next section for guidance)
* also begin with element membership indicator for Excel Names, these also are supposed to clearly show that they are Names, see the respective section below
* elements which are only relevant for development in addition start with dev, i.e. `devfwksExample` would be the codename of a worksheet object belonging to the framework and only required for development
* public procedure names being only relevant for development start with the suffix DEV, e.g. `Public Sub DEV_f_Test()` is a framework sub only relevant for development

### Prefix Indicating Element Membership
Element membership always is explicit in the framework, it is recommended to do this also in the application to reduce cognitive load
* `a`: Application, e.g. `aMUserInterface` is the codename of a module (`M`) belonging to the application
* `f`: Flow Framework 2 (no changes in App Dev) - any changes done in these modules might break the framework and in case of updates also your application, e.g. `fCExample` is the codename of a class module (`C`) belonging to the framework
* `af`: Flow Framework 2 with contents specific to application (only change the concents indicated as changeable), these contents do have to be migrated manually in case of an update, e.g. `afpMGlobals` is the codename of a private module (`pM`) belonging to the framework but designed for holding app specific contents as well

### Object Names in Excel (i.e. not VBA, but as managed via the UI)
* start with element membership prefix: `a`, `f` or `af`
* indicators for Names
	* `named_cell` for Names referring to a range containing only one cell
	* `named_rng` for Names referring to a range containing more than one cell, also including tables
	* `named_fx` for Names referring to a formula
* underscore after membership and Names indicator show that it is an Excel UI and not a VBA element, see examples below in comparison to variable and constant names in VBA
* indicator for Name scope:
	* `g` or omission of scope indicator means reference to range, scoped to workbook ("global")
	* `m` means reference to range, scoped to worksheet ("private")
* rest like in code declarations
* examples
	* `f_named_cell_s_m_VERSION_NUMBER`:
		* `f`: Framework
		* `named_cell`: Excel Name referring to range containing one cell
		* `s`: Represents a string (i.e. cell value to be used as string in VBA)
		* `m`: scoped to worksheet ("private") and thus accessible via `Worksheet.Names(sName)`
		* name in ALL_CAPS: is a constant
	* `a_named_cell_b_g_SayHelloWorld` alias `a_named_cell_b_SayHelloWorld`:
		* `a`: Application
		* `named_cell`: Excel Name referring to range containing one cell
		* `g`: scoped to workbook ("global") and thus accessible via `Workbook.Names(sName)`
		* `b`: Represents a boolean
		* name in CamelCase: is a variable value that might change during usage oof application
### Object Property Names
The intention is to easily see whether it is a property and whether it can be read and/or written to, based on the name alone. Also the type should be clearly indicated.

* `b_prop_rw_NameOfProperty`: a Boolean property for getting and letting
* `s_prop_r_NameOfProperty`: a String property for getting
* `s_prop_w_NameOfProperty`: a String property for letting

### Prefixes for Types
* `b`: Boolean
* `byt`: Byte
* `cur`: Currency
* `d`: Double
* `dte`: Date
* `e`: Enum
* `i`: Integer
* `l`: Long
* `llng`: LongLong
* `lptr`: LongPtr
* `o`: Object
* `oC`: Class Object
* `oCol`: VBA.Collection
* `oDict`: Scripting.Dictionary
* `oFs`: File System Object
* `oWkb`: Workbook Object
* `oWks`: Worksheet Object 
* `oRng`: Range
* `s`: String
* `v`: Variant
* `u`: User Type

* `a` after type prefix: Array, e.g. `vaExample` is the name of an array of type variant when declared with procedure scope

* depending on the scope, underscores are used to more easily identify the scope of a declared name, e.g.
	* `vaExample` for procedure scope
	* `va_m_Example` for module scope (not needing a membership signifier as only accessible within a module)
	* `va_f_p_Example` for project scope, i.e. public in a private module and being part of the framework
	* `va_f_g_Example` for global scope, i.e. public in public module, being part of the framework and also accessible for other VBA projects

### Component Naming
* `wkb`: Workbook
* `wks`: Worksheet
* `F`: user form
* `M`: module
* `C`: class module
* `I`: interface class module
* `p`: private module

* `fFName`: user form with scope `f`, i.e. part of framework
* `afwkbName`: Workbook object with scope `af`, i.e. part of framework but with partly custom contents
* `fwksName`: Worksheet object with scope `f`, i.e. part of framework
* `devfpMName`: Private(`p`) module(`M`) with scope Framework(`f`) and only required when developing(`dev`)
* `aMUserInterface`: Public module with the name UserInterface with scope `a`
* `aCName`: Class module with scope `a`, i.e. part of application
* `IName`: Interface class module with implied scope `a`, i.e. part of application without using the prefix in the name (I don't do it, but it also would work)

### Procedure Naming
Procedure names should indicate the scope, class methods are exempted because the class already indicates the scope, e.g.
* `DEV_f_g_Name`: Public Sub in a Public Module, thus `g` for global, only relevant when developing, thus `DEV` as prefix
* `s_f_Name`: Public Function in a public module returning a string, part of framework - global access implied by not explicitly stating the scope
* `sName`: Public Function in a class module, scope is indicated by class
* `s_f_p_Name`: Public Function in a private module returning a string, part of framework
* `DEV_s_f_m_Name`: Private Function in a private or public module, only relevant for development, returning a string and part of the framework
* `mName`: Private Sub in a class module

## Component Usage and Specific Components
The name prefixes of components are used in the code, especially the ones marking dev contents(`dev`), framework contents(`f`) and contents with framework structure and application contents(`af`) - not usind these prefixes or using them inadequately will break the framework - this is true for all possible types which can be part of the project explorer.

### Workbooks
* `afwkbMain`: The main workbook of the application

### Worksheets
The list contains only the codename of the sheets 
* `afwksErrorLog`: Error log, filled automatically, emptied manually
* `afwksSettings`: The worksheet with the app-specific framework settings
* `awksMain`: One sheet of the application, can be removed if there is at least one other application scope sheet - otherwise switching off maintenance and development mode would fail due to having no worksheet left to be displayed
* `devafwksDevLog`: Worksheet for a development log directly in the workbook, if needed
* `devfwksTestCanvas`: worksheet for development tests of the framework 
* `fwksSettings`: The worksheet with the framework settings


### Modules
* `afpMErrorHandling`: app-specific framework error handling, i.e. custom Enum values and descriptions that can be used in the framework's error handling logic
* `afpMGlobals`: app-specific globals being part of the framework, i.e. custom processing mode for StartProcessing and EndProcessing
* `devfMUserInterface`: user interaction during development
* `devfpMGlobals`: framework globals relevant for development
* `devfpMSandBox`: framework sandbox module relevant for development
* `devfpMTesting`: framework module for running the unit and integration tests during development
* `devfpMUtilities`: utilities for development that do require the frameworks development resources
* `fpMEntryLevel`: framework entry level procedures - the entry level takes care of globals initialization, protection, screen updating etc. - it is a no-brainer wrapper for lower level processing
* `fpMErrorHandling`: framework error handling
* `fpMGlobalsCore`: The module with the framework core globals
* `fpMTemplatesCore`: Template procedures for entry level and non-trivial lower level (i.e. with error handling, testing etc.)
* `fpMUtilities`: framework general utility procedures
* `fpMUtilitiesDev`: framework general utility procedures for development, which must not be removed when deploying - thus "Dev" in this case is not the prefix of the name

### Class Modules
* `devfCUnitTest`: class for unit tests for one unit, relevant for dev only
* `fCCallParams`: class storing information on running procedures, required for error handling, testing etc - mostly meta data which otherwise would not be available, like procedure name and name of the parent component of a procedure
* `fCError`: for storing and using error object information, so that these are retained for proper handling throughout the whole call stack
* `fCRangeArrayProcessor`: convenient array-based range data processing
* `fCSettings`: The class with the framework settings

## Architectural Approach
Code is supposed to be only in forms, modules and class modules, i.e. the workbook, the worksheets and other objects visible in the Microsoft Excel Object tree view in the VBE should not contain any code. The reason are potential severe issues that might occur otherwise, leading to workbook instances broken beyond repair - in such cases, the affected workbooks can't even be opened without Excel crashing.

The overall approach of this framework has three layers:
* public UserInterface modules, the entry point for any user-triggered code execution, usually just calling a sub in an EntryLevel module
* private EntryLevel modules, being the point of entry for code execution, either triggered by a call from a UserInterface module or from an EventHandler. Everything related to sheet protection, deactivating screen processing initializing globals etc. takes place on this entry level. The subs on this level are called entry level subs.
* lower level modules and class modules: the rest of any call stack consists of what in the framework is called "lower level procedures"

There are two procedure types for lower level procedures:
* non-trivial procedures: these might potentially be the place of an error and thus (or for other good reasons) should participate in the error handling logic of the framework and these also can participate in the automated testing - their overall structure is always the same, consisting of a header and declarations section, one or more `try:` sections, one or more `catch:` sections and one or more `finally:` sections.
* trivial procedures: these are so basic that they do not need to participate in the error handling logic of the template - they might have a basic error handling, e.g. just exiting execution with a function's default value in case of an error etc. 
