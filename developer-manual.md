# Developer Manual for Flow Framework 2
The purpose of the framework is to provide a powerful structure for developing and maintaining applications in Excel VBA, which is designed for both clean architecture and clean code.

Much of the documentation you can find directly in the code.

## Naming Conventions
The general purpose of the naming convention is increased readability and reduced cognitive load when developing or reviewing the code. The initial effort for learning the conventions quickly is hugely overcompensated by easy and fast comprehension.

### General Syntax
* the type of variables has to be indicated by a prefix, to reduce cognitive load for developers and to avoid type mismatches
* the nomenclature is based on the so called "Hungarian Notation", see the respective section below for a comprehensive list
* prefixes are in lower case, e.g. s for String or sa for a String array.
* variable names are in CamelCase, e.g. sMyName.
* constant names are in `UPPER_CASE_SEPARATED_BY_UNDERSCORE`, also using a lower case prefix as type indication, e.g. sMY_NAME.
* always begin with the element membership indicator for `Codename` properties of VBA objects, i.e. workbooks, worksheets, modules, class modules (see next section for guidance), e.g. a_wks_Name for a worksheet 
* also begin with the element membership indicator for Excel Names, these also are supposed to clearly show that they are Names, see the respective section below, e.g. s_named_cell_TheNameOfTheCell for a cell the value of which is supposed to be treated as a string
* elements which are only relevant for development in addition start with DEV, e.g. `DEV_f_wks_Example` would be the codename of a worksheet object belonging to the framework and only required for development
* public procedure names being only relevant for development also start with the suffix DEV, e.g. `Public Sub DEV_f_Test()` is a framework sub only relevant for development

### Prefix Indicating Element Membership
Element membership always is explicit in the framework and uses underscores for easier recognition, it is recommended to do this also in the application to reduce cognitive load
* `a`: Application, e.g. `a_M_UserInterface` is the codename of a module (`M`) belonging to the application
* `f`: Flow Framework 2, i.e. no changes in App Dev to these elements as any changes done in these elements might break the framework and in case of updates also your application, e.g. `f_C_Example` is the codename of a class module (`C`) belonging to the framework
* `af`: Flow Framework 2 with contents specific to application. Only change the concents indicated as changeable via code comments. Your app contents do have to be migrated manually in case of an update, which will only deliver the framework portion of it in the first place. Example: `af_pM_Globals` is the codename of a private module (`pM`) belonging to the framework but designed for holding app specific contents as well

### Defined Names in Excel (i.e. not VBA, but as managed via the UI)
* start with element membership indicator prefix: `a`, `f` or `af`
* indicators for Names
	* `named_cell` for Names referring to a range containing only one cell
	* `named_rng` for Names referring to a range containing more than one cell
	* `named_lo` for List Objects
	* `named_fx` for Names referring to a formula
* underscore after membership and Names indicator show that it is an Excel UI and not a VBA element, see examples below in comparison to variable and constant names in VBA
* indicator for Name scope:
	* `g` means reference to range, scoped to workbook ("global")
	* `m` means reference to range, scoped to worksheet ("private")
* rest like in code declarations
* examples
	* `f_named_cell_s_m_VERSION_NUMBER`:
		* `f`: Framework
		* `named_cell`: Excel Name referring to range containing one cell
		* `s`: Represents a string (i.e. cell value to be used as string in VBA)
		* `m`: scoped to worksheet ("private") and thus accessible via `Worksheet.Names(sName)`
		* name in ALL_CAPS: is a constant
	* `a_named_cell_b_g_SayHelloWorld` alias  alias :
		* `a`: Application
		* `named_cell`: Excel Name referring to range containing one cell
		* `g`: scoped to workbook ("global") and thus accessible via `Workbook.Names(sName)`
		* `b`: Represents a boolean
		* name in CamelCase: is a variable value that might change during usage oof application

### Class Property Names
The intention is to easily see whether it is a property and whether it can be read and/or written to, based on the name alone. Also the type should be clearly indicated.

* `b_prop_rw_NameOfProperty`: a Boolean property for getting and letting
* `s_prop_r_NameOfProperty`: a String property for getting
* `s_prop_w_NameOfProperty`: a String property for letting

If a property is private, the respective indicator `m` should be added as the second element of the name, e.g. `b_m_prop_r_NameOfProperty` would be a private read-only property.

### Prefixes for Types
The type prefixes are based on the so called Hungarian notation. You may or may not use underscores between the type prefix and the name of a variable for better readability. If you also need a scope indicator like `m` for "module scope" (or "private"), then you should use underscores. See the examples below.

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
	* `vaExample` or `va_Example` for procedure scope 
	* `va_m_Example` for module scope (not needing a membership signifier as only accessible within a module)
	* `va_f_p_Example` for project scope, i.e. public in a private module and being part of the framework
	* `va_f_g_Example` for global scope, i.e. public in public module, being part of the framework and also accessible for other VBA projects
	* `va_p_Example` or `va_a_p_Example` for project scope of a variable belonging to your application

### Component Naming
Scope indicators:
* `a` or `af` or `f` or `i`: almost like for variables and defined names, in position 1 of the name (unless it is a DEV component), with additional option `i` for independent components, i.e. components that can be used independently from the framework and the application.
* `DEV`: component for development which is not needed in production (always in position 1 of the name)

Component types:
* `wkb`: Workbook
* `wks`: Worksheet
* `F`: user form
* `M`: public module
* `pM`: private module
* `C`: class module
* `I`: interface class module

Examples:
* `f_F_Name`: user form with scope `f`, i.e. part of framework
* `af_wkb_Name`: Workbook object with scope `af`, i.e. part of framework but with partly custom contents
* `f_wks_Name`: Worksheet object with scope `f`, i.e. part of framework
* `DEV_f_pM_Name`: Private(`p`) module(`M`) with scope Framework(`f`) and only required when developing(`DEV`)
* `a_M_UserInterface`: Public module with the name UserInterface with scope `a`
* `a_C_Name`: Class module with scope `a`, i.e. part of application
* `i_I_Name`: Interface class module that can be used independently, i.e. without the framework and also in another project without any need for changes

### Procedure Naming
Procedure names should indicate the scope, i.e. its membership and whether it is public or private.

Class methods are exempted except for private methods, because the class name itself already indicates the scope, i.e. when a method or a property is available for a caller, then it is public.

Examples:
* `DEV_f_g_Name`: Public Sub in a public Module, thus `g` for global, i.e. available also for other projects; only relevant when developing, thus `DEV` as prefix.
* `sName`: Public Function in a class module, scope is implied for caller due to it being accessible.
* `s_f_p_Name`: Public Function in a private module returning a string, part of framework.
* `DEV_s_f_m_Name`: Private Function in a private or public module, only relevant for. development, returning a string and part of the framework
* `mName` or `m_Name`: Private Sub in a (class) module.
* `a_p_Name`: Public sub in a private module belonging to the application, i.e. accessible for other code modules of the same project. `a` is used because module procedures can be called without specifying the module name, this is an important difference in comparison to class modules.

## Component Usage and Specific Components
The name prefixes of components are used by framework code, especially the ones marking dev contents(`DEV`), framework contents(`f`) and contents with framework structure and application contents(`af`) - not using these prefixes in the component naming or using them inadequately will break the framework - this is true for all possible types which can be part of the project explorer.

This means:
* do not rename any `DEV`, `f`, `af` components or you might break the framework
* do not use `f` or `af` as prefixes in the names of your app components as this might lead to unexpected behavior, especially when updating the framework automatically in your app with the respective functionality.
* do use the `DEV` prefix only if you want components to be removed during automatic deployment with the respective functionality and make sure that their absence does not break your app.

### Workbooks
* `a_wkb_Main`: The main workbook of the application - as there only can be one workbook object in a workbook, it was given this name. This is the workbook that contains the framework code.

### Worksheets
`a`-Scope worksheets are not affected when the framework is updated automatically in an application. 

`af`-Scope worksheets are used by framework code, but normally are not affected when the framework is updated - if they are affected, the update comes with specific instructions on how to update them.

`DEV`-Scope worksheets are removed automatically during deployment. Manual follow-up actions may be required to remove refs to the from application code. Also, no references to such sheets should be contained in any worksheets other than `DEV` sheets or in defined names.

The list contains only the codename of the sheets, the names can be changed as you see fit. 

* `a_wks_Main`: an empty worksheet which comes with a fresh framework version - this can be used as main worksheet or removed if there is at least one other `a` worksheet. It is intended to act as the "home page" of an application.
* `a_wks_Settings`: a worksheet which is intended to store the application's settings. It comes with a fresh version of the framework, already containing some basic settings and it is coupled with the class module a_C_Settings, which also comes with a fresh version of the framework.
* `af_wks_ErrorLog`: Error log, filled automatically, emptied manually.
* `af_wks_Settings`: The worksheet with the app-specific framework settings. This means, while these settings are part of the framework, the values stored in this worksheet will be retained when the framework is updated (unless otherwise specified in the update instructions) in your application. Do not change anything in this sheet.
* `DEV_a_wks_TestCanvas`: worksheet for development tests during app development
* `DEF_af_wks_DevLog`: Worksheet for a development log directly in the workbook, if needed. Do not change the column structure in this sheet.
* `DEV_f_wks_TestCanvas`: worksheet for development tests of the framework code
* `f_wks_Settings`: The worksheet with the framework settings. Do not change anything in this sheet. 

### Modules
* `a_M_UserInterface`: public module for app subs and functions which are called directly by user interaction. Normally these just call so called-entry level procedures, please refer to the chapter explaining the recommended architecture.
* `a_pM_EntryLevel`: private module for app entry-level subs and function which are called by code in a user interface module such as `a_M_UserInterface`. Please refer to the chapter explaining the recommended architecture to learn more about entry-level subs and functions.
* `a_pM_Globals`: private module for the application's project scope globals (scope indicator `p` in names), i.e. constants, enumerations, variables and procedures related to project scope global enumerations and variables.
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

## Recommended Tools
I work with following tools for a better dev experience:
* VBE add-in MZ-Tools for numerous helpful features
* VBE add-in RubberduckVBA for testing, refactoring and improved project navigation
* Git and a suitable desktop client, e.g. GitHub Desktop or Sourcetree, for version control
* Mermaid Online Editor for UML class diagrams
* ChatGPT

The following sections contain guidance on how you can use these tools together with the framework.

### Tips for MZ-Tools

### Tips for RubberduckVBA

### Tips for Version Control

#### When working alone without branches

#### When collaborating with branches and folders
The golden rule: no changes whatsoever other than mere code changes when not working on main, i.e. no changes to worksheet contents, no changes to defined names, no new worksheets, no renaming of worksheets etc. 
The reason is simple: git can't handle these and you are very likely to end up with a total mess.
The exception: if you are sure that your changes are clearly separated from the main version of the workbook and can be easily done to the main workbook during integration. In this case the actions for integration should be documented in detail in parallel to your work. They have to be accurate and complete.

### Tips for UML class diagrams made with Mermaid

## Tips for Using ChatGPT
