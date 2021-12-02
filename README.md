# Chairman's Tools

This is the macro collection for the Chairman's Toolbar.

![image](https://user-images.githubusercontent.com/1605764/138600350-0653f21f-7593-43e1-be86-861e5cdd8b0e.png)

These macros were used to help the chair and vice chair of the IFA working group in ETSI ISG NFV standardization org.
The macros assume that the agenda document contains tables of submitted contributions; the exact layout is critical:

              | UID | Title | Allocation | Source | Abstract| Status | Notes |

NOTE: No other tables with exactly 7 columns should be used in the agenda, otherwise some of the macros will fail.

## Workflow
The supported workflow is as follows:
- separate excel macro is used to wrangle the contributions.xls file from https://portal.etsi.org to match the above layout.
- each contribution starts with an empty status, with white/no-color background.
- working through the agenda, each contribution must be given a fixed status. This also helps tracking what is expected to happen next and to guarantee that each contirbution gets handled. 
- with the TODO-tool we can make sure that no contribution remains in an unfinished status. By the end of meeting, each contribution must be with a status of Approved, Noted, Withdrawn or Postponed.
- to speed up the processing of contributions, the linking tool allows quick access to each contribution. The linking tool is also useful when approving agreed documents - if a link exists then the file is available. This avoids the situation when a document has been agreed with a few pending minor fixes and gets approved even when the file of the final version is not yet available.

Note: there are some leftover macros that are not visible in the toolbar, e.g. "email approval" and "further discussion required"

## Features

| Feature description  | Sreenshot |
| ------------- | ------------- |
| Assign status for each contribution, e.g. approved, agreed, revised, noted, postponed, withdrawn - or given a temporary flag of "return" (useful with the TODO-tool, see below). | ![image](https://user-images.githubusercontent.com/1605764/138601520-7de859c4-7278-4a6c-bc7a-ca9a61bae84b.png) |
| Automatically number revisions and create new rows for each revision (prepopulating the data fields from the parent contribution).  | ![image](https://user-images.githubusercontent.com/1605764/138600750-f8ab8eb0-05b4-4943-82a8-a8626a89ccb3.png)  |
|Create hotlinks to files for quick document access. The contribution folder is user definable and the code also supports hard-coded users - based on the OSs user-ID - for easy access (to enable this, set the "Use Default Dir" checkbox).|![image](https://user-images.githubusercontent.com/1605764/138601579-2f279ed4-75bd-4dd6-8b28-5edd3f74413b.png) |
|Calculate statistics on document dispositions and remaining work. This is mainly to estimate the realistic per-document handling time for the remaining contributions.|![image](https://user-images.githubusercontent.com/1605764/138601618-84246fc3-2319-4246-ad9f-f0f135aa47c3.png) |
|Parse current state of the agenda and highlight items that need further work. This can be further limited to documents that are available (i.e. corresponding file exist in the document folder).|![image](https://user-images.githubusercontent.com/1605764/138601682-60a2053f-414f-4687-a1c9-0fc2450b7648.png)|
|Popup to highlight current document ID (helps people at the back rows of large meeting spaces).|![image](https://user-images.githubusercontent.com/1605764/138600914-95c6c5ba-6147-4d8e-95cb-dee76f6b2c8b.png)|




## Known bugs:
- the file-unlinking macro is buggy, haven't figured out why some of the links in the UID column are not cleared away.
- the tabulation of the pop-up window and printed statistics is not always neatly lined up.
- there is not nearly enough error catching in the macros, e.g. the statistics tools get really unhappy if you have other tables
  with 7 rows that don't conform to the layout listed above.

## TODO:
- combine "agreed", "almost feat agreed" and "agreed megaCR" functions - there's a lot of copy&pasting done currently between them.
- optimize the TODO-tool as it now iterates all cells, we could just concentrate on the 6th cell of a row.
- LinkFiles and todo-tool has some copy&pasted code between them, we could clean this up.


## Installation
Word macros are to be part of the .dotm file that also contain the ribbon XML data. The XML was created using the "Custom UI editor for Microsoft Office".

* The .dotm file should be placed in the Word startup folder, typically residing at `C:\Users\USERNAME\AppData\Roaming\Microsoft\Word\STARTUP`
* The .xlam file should be placed in the Excel AddIns folder, typically residing at `C:\Users\USERNAME\AppData\Roaming\Microsoft\AddIns`

A ready package is available as a release (see sidebar).
