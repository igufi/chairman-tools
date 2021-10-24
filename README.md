# Chairman's Tools

This is the macro collection for the Chairman's Toolbar.

These macros were used to help the chair and vice chair of the IFA working group in ETSI ISG NFV standardization org.
The macros assume that the agenda document contains tables of submitted contributions; the exact layout is critical:

              | UID | Title | Allocation | Source | Abstract| Status | Notes |

NOTE: No other tables with exactly 7 columns should be used in the agenda, otherwise some of the macros will fail.

## Workflow
The supported workflow is as follows:
- separate excel macro is used to wrangle the contributions.xls file from https://portal.etsi.org to match the above layout.
- each contribution starts with an empty status, with white/no-color background.
- contributions can be either approved, (feat)agreed, almost feat agreed, revised, noted, postponed, withdrawn or given a temporary
  flag of "return".
- the TODO-tool assumes we want the agenda to only contain contributions with a status of approved, noted, withdrawn or postponed
  and highlights all contributions that do not match those states.
- the linking tool is useful when approving agreed documents - if a link exists then the file is available.

There are some leftover macros that are not visible in the toolbar, e.g. "email approval" and "further discussion required"

The file-linking macro uses the Windows username (hardcoded as per the current chair & vice chair) to quickly populate the correct folder for contributions, for others it will ask the user to point it to the right directory.

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
