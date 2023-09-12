# Python C-Sharp SP/Table Extracter

## What does it do?
Program to extract mentioned SPs and Tables and generate an excel report from
that.

## How does it function?
- User is asked to provide either a file or a folder containing C# Files
- All nested C# files or the single file provided is checked.
- All comments are ignored.
- Then methods for SP and Tables/Inline queries are checked for. These methods are
hardcoded in my code, unfortunately.
- If found then entry is added to the SP column or the Table column
- Some queries are long and have multiple tables mentioned in them. To
tackle this issue, all the table names are extracted using regex pattern matching
and appended to the table list.
- The excel made is formatted automatically with bold headers and borders all around data.
- The text is wrapped and column sensible column width is provided automatically.
- The Excel file generated is timestamped to prevent overwriting.

## Should I use it?
Not meant to be used anywhere. Made for my personal learning purpose.
