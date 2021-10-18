# Overview

Report generator for docx format.
Reads docx template and replaces tokens based on provided logic.

Sample report from old generator see berichtgenerator/
* bericht.doc - template
* preview - generated file
* w97data.txt - commands
* word_makro.txt - makro used for generation

## Features

Available functionality and features are
* Replace tag with text
* Replace tag with table
* Extend existing table when replacing tag in table
* Define template area by start and end tag
* Generate content from template multiple times with different sets of data

## Rules

Text replace
* Replace also within header and footer
* Replace within paragraph with multiple words
* Replace tag in paragraph while preserving style

Table replace
* Only replaces tags within their own paragraphs (not within paragraphs that contain multiple words)
* Replace tag that is not within a table by generating a new table structure
* Replace tag within a table by appending to the table and removing the row with the tag including all other data in that row)