Excel-DNA Custom Registration
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
