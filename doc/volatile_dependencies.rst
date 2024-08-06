Volatile Dependencies
=====================

.. testsetup:: volatile_deps

   import os
   os.chdir(os.path.join("..", "data"))


Support for the Volatile Dependencies part is provided. The resulting object is accessible
as part of the workbook object. It's created on Excel documents that typically feature real
time data or make use of CUBE and it's related functions.

Typical use of the library does not require you to access or modify these elements but should
you require it they can be found as follows:

.. doctest:: volatile_deps

    >>> from openpyxl import load_workbook
    >>>
    >>> wb = load_workbook("sample_with_metadata.xlsx") # This doc uses =CUBE functions
    >>> deps = wb._volatile_deps
    >>> 
    >>> deps.volType[0].type
    'olapFunctions'


.. testcleanup:: volatile_deps

   import os
   os.chdir(os.path.join("..", "tmp"))
