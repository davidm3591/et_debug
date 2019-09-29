# Testing & Debug - Next Steps
## Current Release: 0.4.2-alpha
## Current Bugs

### 1. Local Disk (C:\projects\extraction_tool_project\testing debug files\extraction-testing-09202019\duplicate-errors-ppts) - Release v0.4.2-alpha

| File                                  | Pre, Post-fix (Pass/Fail) | Status         | Notes/Comments          | Validated/Date/Pass/Fail |
|---------------------------------------|---------------------------|----------------|-------------------------|--------------------------|
| 3112-11-12.pptx                       | Pre: Fail                 | Working        | Fail/Pass - Pulls wrong words for vocab and misses real vocab words |     |
|                                       | Pre: Fail                 | Working        | Fail/Pass - empty slides at end are printing |     |
|                                       | Pre: Fail                 | Working        | Fail/Pass - equation codec error |     |
|                                       | Post:                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.|  |
| 3112-12-02.pptx                       | Pre: Fail                 | Working        | Fail/Pass - Misses all vocab words |     |
|                                       | Pre: Fail                 | Working        | Fail/Pass - Multiple blank rows on video <br>qa sheet before Total Time and 'status' legend |     |
|                                       | Pre: Fail                 | Working        | Fail/Pass - empty slides at end are printing |     |
|                                       | Pre: Fail                 | Working        | Fail/Pass - equation codec error |     |
|                                       | Post:                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.|  |
| 3112-12-04.pptx                       | Pre: Fail                 | Working        | Fail/Pass - Misses all vocab words |                          |
|                                       | Pre: Fail                 | Working        | Fail/Pass - Multiple blank rows on video <br>qa sheet before Total Time and 'status' legend |     |
|                                       | Pre: Fail                 | Working        | Fail/Pass - empty slides at end are printing |     |
| 8525-03-02.pptx                       | Pre: Fail                 | Working        | Fail/Pass - Pulls wrong words for vocab and misses real vocab words |     |
|                                       | Pre: Fail                 | Working        | Fail/Pass - Multiple blank rows on video <br>qa sheet before Total Time and 'status' legend |     |
|                                       | Pre: Fail                 | Working        | Fail/Pass - empty slides at end are printing |     |
| 8525-01-06.pptx                       | Pre: Fail                 | Working        | Fail/Pass - Missing data, line 15 (slide 9) |                          |
|                                       | Pre: Fail                 | Working        | Fail/Pass - Multiple blank rows on video <br>qa sheet before Total Time and 'status' legend |     |
|                                       | Pre: Fail                 | Working        | Fail/Pass - Pulls wrong words for vocab and misses real vocab words |     |
|                                       | Pre: Fail                 | Working        | Fail/Pass - empty slides at end are printing |     |
| 8525-04-06.pptx                       | Pre: Fail                 | Working        | Fail/Pass - Missing data, line 15 (slide 9) |                          |

### 2. Local Disk (C:\projects\extraction_tool_project\testing_debug_files\extraction-testing-09102019\failed-ppts) - Release: v0.4.1-alpha

| File                    | Pre, Post-fix (Pass/Fail)      | Status         | Notes/Comments          | Validated/Date/Pass/Fail                 |
|-------------------------|--------------------------------|----------------|-------------------------|------------------------------------------|
| 3112-09-02.pptx         | Pre: Fail                      | Working        | Fail/Pass - equation codec error |    |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
| 3112-09-06.pptx         | Pre: Fail                      | Working        | Fail/Pass - equation codec error          |  |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
| 3112-09-10.pptx         | Pre: Fail                      | Working        | Fail/Pass - equation codec error |    |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
| 3112-10-04.pptx         | Pre: Fail                      | Working        | Fail/Pass - equation codec error |    |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
| 3112-10-06.pptx         | Pre: Fail                      | Working        | Fail/Pass - equation codec error |    |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |
| 8103-11-04.pptx         | Pre: Fail                      | Working        | Fail/Pass - equation codec error |    |
|                         | Post: Pass                     | Fixed-complete | Temp Solution-Output Equation Error File.<br>User receives (1)"Text Format Error" pop-up warning and<br>(2)"Frame Content Waring" prints to `output` directory.| 09/26/19, Pass, SW |

______


## Completed Bugs

### No Vocab (C:\projects\extraction_tool_project\testing_debug_files\no_vocab_ppts) - Release: v0.4.0-alpha

__Part 1__

| File                                  | Pre, Post-fix (Pass/Fail)      | Status         | Notes/Comments          | Validated/Date/Pass/Fail |
|---------------------------------------|--------------------------------|----------------|-------------------------|--------------------------|
| 8309-02-15-Lab_Cell Size.pptx         | Post: Pass                     | Fixed-complete | PPT has no vocab words  | 09/24/19, Pass, SW |
| 8309-05-07-Mendelian Inheritance.pptx | Post: Pass                     | Fixed-complete | PPT has no vocab words  | 09/25/19, Pass, SW |
| 8309-07-03-Natural Selection.pptx     | Post: Pass                     | Fixed-complete | PPT has no vocab words  | 09/25/19, Pass, SW |
| 8309-07-04-Artificial Selection.pptx  | Post: Pass                     | Fixed-complete | PPT has no vocab words  | 09/24/19, Pass, SW |
| 8309-07-12-Common Ancestry.pptx       | Post: Pass                     | Fixed-complete | PPT has no vocab words  | 09/24/19, Pass, SW |


### No Vocab (C:\projects\extraction_tool_project\testing_debug_files\no_vocab_ppts) - Release: v0.4.1-alpha

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


### extraction-testing-09102019 (C:\projects\extraction_tool_project\testing_debug_files\extraction-testing-09102019\pass-fail-ppts) - Release: v0.4.0-alpha

| File             | Pre, Post-fix (Pass/Fail) | Status         | Notes/Comments                                           | Validated/Date/Pass/Fail |
|------------------|---------------------------|----------------|----------------------------------------------------------|--------------------------|
|8309-02-15.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing data - task frames that have no audio     | 09/24/19, Pass, SW |
|8309-07-15.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing "Activity Type"                           | 09/24/19, Pass, SW |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 49 Slide Layout should not show | 09/24/19, Pass, SW |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 49 Filename is missing          | 09/24/19, Pass, SW |
|8525-01-01.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing data - task frames that have no audio     | 09/24/19, Pass, SW |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 67 Slide Layout should not show | 09/24/19, Pass, SW |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 67 Filename is missing          | 09/24/19, Pass, SW |
|8525-01-04.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing data - task frames that have no audio     | 09/24/19, Pass, SW |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 64 Slide Layout should not show | 09/24/19, Pass, SW |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 64 Filename is missing          | 09/24/19, Pass, SW |
|8705-07-04.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing data - task frames that have no audio     | 09/24/19, Pass, SW |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 58 Slide Layout should not show | 09/24/19, Pass, SW |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 58 Filename is missing          | 09/24/19, Pass, SW |
|8705-07-08.pptx   | Post: Pass                | Fixed-complete | FIXED: Missing data - task frames that have no audio     | 09/24/19, Pass, SW |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 68 Slide Layout should not show | 09/24/19, Pass, SW |
|                  | Post: Pass                | Fixed-complete | FIXED: No error msg-line 68 Filename is missing          | 09/24/19, Pass, SW |
