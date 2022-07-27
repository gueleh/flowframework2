# Developer Manual for Flow Framework 2
Most of the documentation you can find directly in the code.

## Naming Conventions
### General Syntax
* prefixes are in lower case
* variable names are in CamelCase
* constants are in UPPER_CASE_SEPARATED_BY_UNDERSCORE
* always begin with membership indicator for name properties of VBA objects, i.e. workbooks, worksheets, modules, class modules
* also begin with membership indicator for Excel names

### Prefix Indicating Element Membership
Element membership always is explicit in the framework, it is recommended to do this also in the application to reduce cognitive load
* a: Application
* f: Flow Framework (no changes in App Dev) - any changes done in these modules might break the framework and in case of updates also your application
* af: Flow Framework with contents specific to application (only change the concents indicated as changeable), these contents do have to be migrated manually in case of an update

### Object Names in Excel (i.e. not VBA)
* start with element membership prefix: a/f/af
* indicator for Name: n
* underscore after membership and name indicator to indicate that it is an Excel and not a VBA object, see examples below in comparison to variable and constant names in VBA
* indicator for Name scope:
	* g means reference to range, scoped to workbook ("global")
	* m means reference to range, scoped to worksheet ("private")
	* x means that it is a named formula
* rest like in code declarations
* examples
	* fn_smVERSION_NUMBER:
		* f: Framework
		* n: Excel Name
		* m: Reference to range, scoped to worksheet ("private")
		* s: Represents a string
		* name in all caps: is a constant
	* an_bgSayHelloWorld:
		* a: Application
		* n: Excel Name
		* g: Reference to range, scoped to workbook ("global")
		* b: Represents a boolean
		* name in CamelCase: is a variable value that might change during usage oof application

### Prefixes for Types
* b: Boolean
* byt: Byte
* cur: Currency
* d: Double
* dte: Date
* e: Enum
* i: Integer
* l: Long
* llng: LongLong
* lptr: LongPtr
* o: Object
* oC: Class Object
* oCol: VBA.Collection
* oDict: Scripting.Dictionary
* oFs: File System Object
* oWkb: Workbook Object
* oWks: Worksheet Object 
* rng: Range
* s: String
* v: Variant
* u: User Type

* a after type prefix: Array, e.g. array of type variant starts with prefix va

* M: module
* C: class module
* I: interface class module

## Component Usage and Specific Components
### Workbooks
* afwkbMain: The main workbook of the application

### Worksheets
The list contains only the codename of the sheets 
* afwksSettings: The worksheet with the app-specific framework settings
* fwksSettings: The worksheet with the framework settings

### Modules By Functional Units
* afpMErrorHandling
* fpMErrorHandling

* fpMGlobalsCore: The module with the framework core globals

### Class Modules
* fCSettings: The class with the framework settings

## Architectural Approach
Code is supposed to be only in forms, modules and class modules, i.e. the workbook, the worksheets and other objects visible in the Microsoft Excel Object tree view in the VBE should not contain any code. The reason are potential severe issues that might occur otherwise, leading to workbook instances broken beyond repair - in such cases, the affected workbooks can't even be opened without Excel crashing.

The overall approach of this framework has three layers:
* public UserInterface modules, the entry point for any user-triggered code execution, usually just calling a sub in an EntryLevel module
* private EntryLevel modules, being the point of entry for code execution, either triggered by a call from a UserInterface module or from an EventHandler. Everything related to sheet protection, deactivating screen processing initializing globals etc. takes place on this entry level. The subs on this level are called entry level subs.
* lower level modules and class modules: the rest of any call stack consists of what in the framework is called "lower level procedures"

There are two procedure types for lower level procedures:
* non-trivial procedures: these might potentially be the place of an error and thus (or for other good reasons) should participate in the error handling logic of the framework and these also can participate in the automated testing - their overall structure is always the same, consisting of a header and declarations section, one or more "try:" sections, one or more "catch:" sections and one or more "finally:" sections.
* trivial procedures: these are so basic that they do not need to participate in the error handling logic of the template - they might have a basic error handling, e.g. just exiting execution with a function's default value in case of an error etc. 
