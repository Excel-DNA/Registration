Excel-DNA Registration Helper
=============================

This library implements helper functions to assist and modify the Excel-DNA function registration, by applying various transformations before the functions are registered.

The following transformations have been implemented:

Generation of wrapper functions for:

- Functions returning Task<T> or IObservable<T> as asynchronous or RTD-based functions (including F# Async<T> functions)
- Optional parameters (with default values), 'params' parameters and Nullable<T> parameters
- Range parameters in Visual Basic functions

Examples of general function transformations:

- Logging / Caching / Timing handlers
- Suppress in Function Arguments dialog

_If you've previously used the CustomRegistration library, note that I've renamed and rearranged the project source, and renamed the output assembly from ExcelDna.CustomRegistration to ExcelDna.Registration. The last state of the project before the large-scale rearrangement is marked by the git tag **CustomRegistration_Before_Rename**, and can be retrieved from the release tab on GitHub._
