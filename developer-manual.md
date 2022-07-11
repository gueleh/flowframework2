# Developer Manual for Flow Framework 2
Most of the documentation you can find directly in the code.

## Naming Conventions
### Prefix Indicating Element Membership
* a: Application
* f: Flow Framework (no changes in App Dev)
* af: Flow Framework with contents specific to application (only change the concents indicated as changeable)

### Object Names in Excel
* start with element membership prefix: a/f/af
* indicator for Name: n
* indicator for Name scope:
	* g means reference to range, scoped to workbook ("global")
	* m means reference to range, scoped to worksheet ("private")
	* x means that it is a named formula
* rest like in code declarations
* examples
	* fn_msVERSION_NUMBER:
		* f: Framework
		* n: Excel Name
		* m: Reference to range, scoped to worksheet ("private")
		* s: Represents a string
		* name in all caps: is a constant
	* an_gbSayHelloWorld:
		* a: Application
		* n: Excel Name
		* g: Reference to range, scoped to workbook ("global")
		* b: Represents a boolean
		* name in CamelCase: is a variable value that might change during usage oof application

### Component Usage and Specific Components
#### Workbooks
* afwkbMain: The main workbook of the application

#### Worksheets
* afwksSettings: The worksheet with the app-specific framework settings
* fwksSettings: The worksheet with the framework settings

#### Modules By Functional Units
* afmErrorHandling
* fmErrorHandling

* fmGlobalsCore: The module with the framework core globals

#### Class Modules
* fclsSettings: The class with the framework settings
	