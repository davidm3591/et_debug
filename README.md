# Testing & Debug - Next Steps
## Current Release: 0.4.5-alpha
## Current Bugs



### Codec Error Crash Files - Release: v0.4.4-alpha

| File                | Issue                     | Status         | Notes/Comments                                       | Validated/Date/Pass/Fail |
|---------------------|---------------------------|----------------|------------------------------------------------------|--------------------------|
| __3112-09-02.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __3112-09-06.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __3112-09-10.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __3112-10-04.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __3112-10-06.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __8103-11-04.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __3112-11-12.pptx__ |                           |                | ..\extraction-testing-09202019\duplicate-errors-ppts |                          |
| __3112-12-02.pptx__ |                           |                | ..\extraction-testing-09202019\duplicate-errors-ppts |                          |
<br>

### Timing - Release: v0.4.5-alpha (Will update in future update.)

| File                | Issue                     | Status         | Notes/Comments                                       | Validated/Date/Pass/Fail |
|---------------------|---------------------------|----------------|------------------------------------------------------|--------------------------|
|                     | Timing accepted when written as `Timing: 0.5 min` or `Timing:0.5 min`.  |  |  |  |
<br>
______


## Completed Bugs

### Bug: Missing and wrong time format (not "min") do not throw error and trigger message - Release: v0.4.5-alpha

|File                  | Bug Description                 | Validated/Date/Pass/Fail |
|----------------------|---------------------------------|--------------------------|
| __8525-03-02.pptx__  | If timing unit is blank (missing "min, minute,or minutes")<br>there is no error, and time is not counted in the extraction sheet. | 10/16/19, Pass, SW |
| 8522-04-08 | “ | 10/16/19, Pass, SW |
| 8705-07-01 | “ | 10/16/19, Pass, SW |
| 8702-03-01 | “ | 10/16/19, Pass, SW |
| 8309-05-03 | “ | 10/16/19, Pass, SW |
<br>

### Local Disk (C:\projects\extraction_tool_project\testing_debug_files\ext_testing_workbook_10072019) - Release: v0.4.4-alpha
#### Activities (framechains) starting with non-video frame are not showing in the extraction sheet
| File                | Pre, Post-fix (Pass/Fail) | Status           | Notes/Comments                              | Validated/Date/Pass/Fail                                 |
|---------------------|---------------------------|------------------|---------------------------------------------|----------------------------------------------------------|
| __3112-06-01.pptx__ | Post: Pass                | complete - fixed | Bug does not occur in Office 365 - resolved | Verified by Sam on<br>ext_testing_workbook_10072019.xlsx |
| __3112-06-03.pptx__ | Post: Pass                | complete - fixed | Bug does not occur in Office 365 - resolved | Verified by Sam on<br>ext_testing_workbook_10072019.xlsx |
| __3112-06-05.pptx__ | Post: Pass                | complete - fixed | Bug does not occur in Office 365 - resolved | Verified by Sam on<br>ext_testing_workbook_10072019.xlsx |
| __3112-06-07.pptx__ | Post: Pass                | complete - fixed | Bug does not occur in Office 365 - resolved | Verified by Sam on<br>ext_testing_workbook_10072019.xlsx |
| __3112-06-09.pptx__ | Post: Pass                | complete - fixed | Bug does not occur in Office 365 - resolved | Verified by Sam on<br>ext_testing_workbook_10072019.xlsx |
| __3112-06-11.pptx__ | Post: Pass                | complete - fixed | Bug does not occur in Office 365 - resolved | Verified by Sam on<br>ext_testing_workbook_10072019.xlsx |
| __3112-06-13.pptx__ | Post: Pass                | complete - fixed | Bug does not occur in Office 365 - resolved | Verified by Sam on<br>ext_testing_workbook_10072019.xlsx |
| __3112-06-15.pptx__ | Post: Pass                | complete - fixed | Bug does not occur in Office 365 - resolved | Verified by Sam on<br>ext_testing_workbook_10072019.xlsx |
| __3112-06-17.pptx__ | Post: Pass                | complete - fixed | Bug does not occur in Office 365 - resolved | Verified by Sam on<br>ext_testing_workbook_10072019.xlsx |
| __3112-06-19.pptx__ | Post: Pass                | complete - fixed | Bug does not occur in Office 365 - resolved | Verified by Sam on<br>ext_testing_workbook_10072019.xlsx |
<br>

### Multiple Blank Lines Between Last Video and Status Legend - Release: v0.4.3-alpha (new release with this fix: v0.4.4-alpha)
| File                | Issue                     | Status         | Notes/Comments                                       | Validated/Date/Pass/Fail |
|---------------------|---------------------------|----------------|------------------------------------------------------|--------------------------|
| __3112-11-12.pptx__ | Multiple blank lines      | Fixed-complete | ..\extraction-testing-09202019\duplicate-errors-ppts | 10/06/2019, Pass, drm<br>10/07/2019, Pass, sw |
| __3112-12-02.pptx__ | Multiple blank lines      | Fixed-complete | ..\extraction-testing-09202019\duplicate-errors-ppts | 10/06/2019, Pass, drm<br>10/07/2019, Pass, sw |
| __3112-12-04.pptx__ | Multiple blank lines      | Fixed-complete | ..\extraction-testing-09202019\duplicate-errors-ppts | 10/06/2019, Pass, drm<br>10/07/2019, Pass, sw |
| __8525-01-06.pptx__ | Multiple blank lines      | Fixed-complete | ..\extraction-testing-09202019\duplicate-errors-ppts | 10/06/2019, Pass, drm |
| __8525-03-02.pptx__ | Multiple blank lines      | Fixed-complete | ..\extraction-testing-09202019\duplicate-errors-ppts | 10/06/2019, Pass, drm |
<br>

### Local Disk (C:\projects\extraction_tool_project\testing_debug_files\extraction-testing-09262019\pass-fail-ppts) - Release: v0.4.3-alpha

| File                | Pre, Post-fix (Pass/Fail) | Status         | Notes/Comments                                    | Validated/Date/Pass/Fail |
|---------------------|---------------------------|----------------|---------------------------------------------------|--------------------------|
| __3112-09-10.pptx__ | Post: Pass  | Fixed-complete | Missing vocab   | 10/04/2019, Pass, SW |
| __3112-09-02.pptx__ | Post: Pass  | Fixed-complete | Missing vocab   | 10/04/2019, Pass, SW |
| __3112-09-06.pptx__ | Post: Pass  | Fixed-complete | Missing vocab   | 10/04/2019, Pass, SW |
| __3112-10-06.pptx__ | Post: Pass  | Fixed-complete | Missing vocab   | 10/04/2019, Pass, SW |
| __3112-10-04.pptx__ | Post: Pass  | Fixed-complete | Missing vocab   | 10/04/2019, Pass, SW |
| __3112-11-12.pptx__ | N/A                       | N/A Duplicate  | Missing vocab - Duplicate                         | see ..\extraction-testing-09202019\duplicate-errors-ppts |
| __3112-12-02.pptx__ | N/A                       | N/A Duplicate  | Missing vocab - Duplicate                         | see ..\extraction-testing-09202019\duplicate-errors-ppts |
| __3112-12-04.pptx__ | N/A                       | N/A Duplicate  | Incorrect & missing vocab - Duplicate             | see ..\extraction-testing-09202019\duplicate-errors-ppts |
| __8525-01-06.pptx__ | N/A                       | N/A Duplicate  | Incorrect & missing vocab - Duplicate             | see ..\extraction-testing-09202019\duplicate-errors-ppts |
| __8103-11-04.pptx__ | Post: Pass  | Fixed-complete | Missing vocab   | 10/04/2019, Pass, SW |
| __8525-03-02.pptx__ | N/A                       | N/A Duplicate  | Incorrect & missing vocab - Duplicate             | see ..\extraction-testing-09202019\duplicate-errors-ppts |
<br>

### Local Disk (C:\projects\extraction_tool_project\testing debug files\extraction-testing-09202019\duplicate-errors-ppts) - Release v0.4.2-alpha

| File                                  | Pre, Post-fix (Pass/Fail) | Status         | Notes/Comments          | Validated/Date/Pass/Fail |
|---------------------------------------|---------------------------|----------------|-------------------------|--------------------------|
| __3112-11-12.pptx__                   | Post: Pass                | Fixed-complete | Pulls wrong words for vocab and misses real vocab words.<br>Changed dashes to hyphens | 10/02/2019, Pass, drm |
|                                       | Post: Pass                | Fixed-complete | Empty slides at end are printing.<br>Moved graphics info slide to front info section of PPT | 10/02/2019, Pass, drm |
|                                       | ~~Pre: Fail~~             | ~~Working~~    | ~~Fail/Pass - equation codec error~~ | **Moved to Codec Error Crash Files**   |
|                                       | Post: Pass                | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.|  |
| __3112-12-02.pptx__                   | Post: Pass                | Fixed-complete | Misses all vocab words.<br>Changed dashes to hyphens. | 10/02/2019, Pass, drm |
|                                       | ~~Pre: Fail~~             | ~~Working~~    | ~~Fail/Pass - Multiple blank rows on video <br>qa sheet before Total Time and 'status' legend~~ | **Moved to Multiple Blank Lines...**   |
|                                       | Post: Pass                | Fixed-complete | Empty slides at end are printing.<br>Moved graphics info slide to front info section of PPT | 10/02/2019, Pass, drm |
|                                       | ~~Pre: Fail~~             | ~~Working~~    | ~~Fail/Pass - equation codec error~~ | **Moved to Codec Error Crash Files**   |
|                                       | Post: Pass                | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.|  |
| __3112-12-04.pptx__                   | Post: Pass                | Fixed-complete | Misses all vocab words.<br>Changed dashes to hyphens. | 10/02/2019, Pass, drm |
|                                       | ~~Pre: Fail~~             | ~~Working~~    | ~~Fail/Pass - Multiple blank rows on video <br>qa sheet before Total Time and 'status' legend~~ | **Moved to Multiple Blank Lines...**   |
|                                       | Post: Pass                | Fixed-complete | Empty slides at end are printing.<br>Moved graphics info slide to front info section of PPT | 10/02/2019, Pass, drm |
| __8525-03-02.pptx__                   | Post: Pass                | Fixed-complete | Pulls wrong words for vocab and misses real vocab words.<br>Changed dashes to hyphens  | 10/02/2019, Pass, drm |
|                                       | ~~Pre: Fail~~             | ~~Working~~    | ~~Fail/Pass - Multiple blank rows on video <br>qa sheet before Total Time and 'status' legend~~ | **Moved to Multiple Blank Lines...**   |
|                                       | Post: Pass                | Fixed-complete | Empty slides at end are printing.<br>Moved graphics info slide to front info section of PPT | 10/02/2019, Pass, drm |
| __8525-01-06.pptx__                   | Post: Pass                | Fixed-complete | Missing data, line 15 (slide 9) | 10/01/2019, Pass, SW |
|                                       | ~~Pre: Fail~~             | ~~Working~~    | ~~Fail/Pass - Multiple blank rows on video <br>qa sheet before Total Time and 'status' legend~~ | **Moved to Multiple Blank Lines...**   |
|                                       | Post: Pass                | Fixed-complete | Pulls wrong words for vocab and misses real vocab words.<br>Changed dashes to hyphens | 10/01/2019, Pass, SW |
|                                       | Post: Pass                | Fixed-complete | Empty slides at end are printing.<br>Moved graphics info slide to front info section of PPT | 10/02/2019, Pass, drm |
| __8525-04-06.pptx__                   | Post: Pass                | Fixed-complete | Fail/Pass - Missing data, line 15 (slide 9) | 10/01/2019, Pass, SW |
<br>

### Local Disk (C:\projects\extraction_tool_project\testing_debug_files\extraction-testing-09102019\failed-ppts) - Release: v0.4.1-alpha

| File                    | Pre, Post-fix (Pass/Fail)      | Status         | Notes/Comments          | Validated/Date/Pass/Fail                 |
|-------------------------|--------------------------------|----------------|-------------------------|------------------------------------------|
| __3112-09-02.pptx__     | ~~Pre: Fail~~                      | ~~Working~~        | ~~Fail/Pass - equation codec error~~ | **Moved to Codec Error Crash Files**   |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
| __3112-09-06.pptx__     | ~~Pre: Fail~~                      | ~~Working~~        | ~~Fail/Pass - equation codec error~~          | **Moved to Codec Error Crash Files**   |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
| __3112-09-10.pptx__     | ~~Pre: Fail~~                      | ~~Working~~        | ~~Fail/Pass - equation codec error~~ | **Moved to Codec Error Crash Files**   |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
| __3112-10-04.pptx__     | ~~Pre: Fail~~                      | ~~Working~~        | ~~Fail/Pass - equation codec error~~ | **Moved to Codec Error Crash Files**   |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
| __3112-10-06.pptx__     | ~~Pre: Fail~~                      | ~~Working~~        | ~~Fail/Pass - equation codec error~~ | **Moved to Codec Error Crash Files**   |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
| __8103-11-04.pptx__     | ~~Pre: Fail~~                      | ~~Working~~        | ~~Fail/Pass - equation codec error~~ | **Moved to Codec Error Crash Files**   |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
<br>

### Local Disk (C:\projects\extraction_tool_project\testing_debug_files\test_files_sam_08132019) - Release: v0.4.0-alpha

| File                | Pre, Post-fix (Pass/Fail) | Status   | Notes/Comments                                    | Validated/Date/Pass/Fail |
|---------------------|---------------------------|----------|---------------------------------------------------|--------------------------|
| __8309-02-15.pptx__ |                           | Complete | see ..\extraction-testing-09102019\pass-fail-ppts | N/A                      |
| __8518-06-06.pptx__ |                           |          | **See 8518-06-06-notes.md**                       |                          |
| __8525-01-01.pptx__ |                           | Complete | see ..\extraction-testing-09102019\pass-fail-ppts | N/A                      |
| __8705-07-04.pptx__ |                           | Complete | see ..\extraction-testing-09102019\pass-fail-ppts | N/A                      |
| __3112-09-02.pptx__ |                           | Complete | see ..\extraction-testing-09102019\failed-ppts    | N/A                      |
<br>

### Local Disk (C:\projects\extraction_tool_project\testing_debug_files\no_vocab_ppts) - Release: v0.4.0-alpha

__Part 1__

| File                                      | Pre, Post-fix (Pass/Fail)      | Status         | Notes/Comments          | Validated/Date/Pass/Fail |
|-------------------------------------------|--------------------------------|----------------|-------------------------|--------------------------|
| __8309-02-15-Lab_Cell Size.pptx__         | Post: Pass                     | Fixed-complete | PPT has no vocab words  | 09/24/19, Pass, SW       |
| __8309-05-07-Mendelian Inheritance.pptx__ | Post: Pass                     | Fixed-complete | PPT has no vocab words  | 09/25/19, Pass, SW       |
| __8309-07-03-Natural Selection.pptx__     | Post: Pass                     | Fixed-complete | PPT has no vocab words  | 09/25/19, Pass, SW       |
| __8309-07-04-Artificial Selection.pptx__  | Post: Pass                     | Fixed-complete | PPT has no vocab words  | 09/24/19, Pass, SW       |
| __8309-07-12-Common Ancestry.pptx__       | Post: Pass                     | Fixed-complete | PPT has no vocab words  | 09/24/19, Pass, SW       |
<br>

### NO VOCAB - Local Disk (C:\projects\extraction_tool_project\testing_debug_files\no_vocab_ppts) - Release: v0.4.1-alpha

__Part 2__

| File                                | Pre, Post-fix (Pass/Fail) | Status               | Notes/Comments                              | Validated/Date/Pass/Fail |
|------------------------------------ |---------------------------|----------------------|---------------------------------------------|--------------------------|
| __8526-10-02 Unit 1 Review.pptx__   | Post: Pass                | Fixed-complete       | PPT has no vocab words                      | 09/24/19, Pass, SW |
|                                     | Post: Pass                | Fixed-complete       | No error msg - leading task frame missing*  | 09/25/19, Pass, SW |
| __8526-10-03 Unit 2 Review.pptx__   | Post: Pass                | Fixed-complete       | PPT has no vocab words                      | 09/24/19, Pass, SW |
|                                     | Post: Pass                | Fixed-complete       | No error msg - leading task frame missing   | 09/25/19, Pass, SW |
| __8526-10-04 Unit 3 Review.pptx__   | Post: Pass                | Fixed-complete       | PPT has no vocab words                      | 09/24/19, Pass, SW |
|                                     | Post: Pass                | Fixed-complete       | No error msg - leading task frame missing   | 09/25/19, Pass, SW |
| __8526-10-06 Unit 5 Review.pptx__   | Post: Pass                | Fixed-complete       | PPT has no vocab words                      | 09/24/19, Pass, SW |
|                                     | Post: Pass                | Fixed-complete       | No error msg - leading task frame missing   | 09/25/19, Pass, SW |
______

### Error Message(s)
__*8526-10-02 Unit 1 Review.pptx error message:__
<pre>
Exception in Tkinter callback
Traceback (most recent call last):
  File "c:\users\dmilatz\anaconda3\Lib\tkinter\__init__.py", line 1705, in __call__
    return self.func(*args)
  File ".\extract_build\extraction_pptx.py", line 205, in get_file
    read_slide(filename_path)
  File ".\extract_build\extraction_pptx.py", line 851, in read_slide
    basic_info, basic_vid_qa)
  File ".\extract_build\extraction_out.py", line 939, in excel_data_out
    row, 9, all_act_times[i], ext_sht_content_line)
IndexError: list index out of range
</pre>
<br>

### extraction-testing-09102019 (C:\projects\extraction_tool_project\testing_debug_files\extraction-testing-09102019\pass-fail-ppts) - Release: v0.4.0-alpha

| File             | Pre, Post-fix (Pass/Fail) | Status         | Notes/Comments                                           | Validated/Date/Pass/Fail |
|------------------|---------------------------|----------------|----------------------------------------------------------|--------------------------|
|8309-02-15.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing data - task frames that have no audio     | 09/24/19, Pass, SW       |
|8309-07-15.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing "Activity Type"                           | 09/24/19, Pass, SW       |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 49 Slide Layout should not show | 09/24/19, Pass, SW       |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 49 Filename is missing          | 09/24/19, Pass, SW       |
|8525-01-01.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing data - task frames that have no audio     | 09/24/19, Pass, SW       |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 67 Slide Layout should not show | 09/24/19, Pass, SW       |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 67 Filename is missing          | 09/24/19, Pass, SW       |
|8525-01-04.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing data - task frames that have no audio     | 09/24/19, Pass, SW       |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 64 Slide Layout should not show | 09/24/19, Pass, SW       |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 64 Filename is missing          | 09/24/19, Pass, SW       |
|8705-07-04.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing data - task frames that have no audio     | 09/24/19, Pass, SW       |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 58 Slide Layout should not show | 09/24/19, Pass, SW       |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 58 Filename is missing          | 09/24/19, Pass, SW       |
|8705-07-08.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing data - task frames that have no audio     | 09/24/19, Pass, SW       |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 68 Slide Layout should not show | 09/24/19, Pass, SW       |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 68 Filename is missing          | 09/24/19, Pass, SW       |
