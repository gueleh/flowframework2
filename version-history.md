# Version History of Flow Framework 2
## 1.0.2 230620
* fixed entry level template for more transparency regarding start and end processing

## 1.0.1 230616
* fixed nomenclature issue for some named cells and respective contents
* updated developer manual

## 1.0.0 230616
* reduced cognitive load for dev work: changed names of classes and modules to be more easily readable, losing direct compatibility with prior versions
* added containers for APP content
* added features from FF'2 Little Sis which were not in before

## 0.15.0, 230415
* added f_C_Wks for data worksheet handling
* added f_I_DataRecord and f_C_DataRecord for data record handling

## 0.14.0, 230131
* added export of wks names and codenames to textfile for version control
* added export of vb project code library references to textfile for version control

## 0.13.0, 230131
* added dev utility to send a ping to the direct window informing about module and proc name, in order to manually check whether logic skeletons are properly integrated in the call stack

## 0.12.1, 230131
* bugfix: added Option Private Module to devfpMUtilities

## 0.12.0, 220816
* export of all code modules via VBA, including the worksheet .cls files
* export the properties of all Name objects except for the value to a text file for version control

## 0.11.0, 220805
* added skeleton for convenient array-based range data processing class, containing method to sanitize array items starting with a 0, so that the leading zeroes are retained
* added test canvas worksheet for framework development
* added dev utility to reset test canvas worksheet

## 0.10.0, 220804
* refactored and updated dev manual

## 0.9.0, 220803
* added Development Mode

## 0.8.0, 220803
* set visibility of technical Names automatically

## 0.7.0, 220802
* refactoring, updated dev manual

## 0.6.0, 220727
* refactored code base, see commit changes

## 0.5.0, 220715
* refactored code base, see commit changes

## 0.4.0, 220715
* refactored, mostly changing names to new syntax allowing MZ Tools code review

## 0.3.0, 220715
* refactored, mostly changing names to new syntax allowing MZ Tools code review

## 0.2.0, 220711
* added dev log and basic sub to quickly set items to done

## 0.1.0, 220709
* initial creation containing the basics of a template