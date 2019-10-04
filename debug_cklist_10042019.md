# ET Debug Checklist

## Bug: Local Disk (C:\projects\extraction_tool_project\testing debug files\extraction-testing-09202019\duplicate-errors-ppts) - Release v0.4.2-alpha  

## Bug: Multiple Blank Lines Between Last Video and Status Legend - Release: v0.4.3-alpha

* [ ] 1. Review tag type code and output video code
* [ ] 2. Check method to catch blank fields on Video QA output
 - (Note: can't use new tag without recoding all of the video tag type code)

| File                | Pre, Post-fix (Pass/Fail) | Status         | Notes/Comments                                       | Validated/Date/Pass/Fail |
|---------------------|---------------------------|----------------|------------------------------------------------------|--------------------------|
| __3112-12-04.pptx__ |                           |                | ..\extraction-testing-09202019\duplicate-errors-ppts |                          |
| __8525-03-02.pptx__ |                           |                | ..\extraction-testing-09202019\duplicate-errors-ppts |                          |
| __8525-01-06.pptx__ |                           |                | ..\extraction-testing-09202019\duplicate-errors-ppts |                          |
| __3112-12-02.pptx__ |                           |                | ..\extraction-testing-09202019\duplicate-errors-ppts |                          |




## Bug: Codec Error Crash Files - Release: v0.4.3-alpha

| File                | Pre, Post-fix (Pass/Fail) | Status         | Notes/Comments                                       | Validated/Date/Pass/Fail |
|---------------------|---------------------------|----------------|------------------------------------------------------|--------------------------|
| __3112-09-02.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __3112-09-06.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __3112-09-10.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __3112-10-04.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __3112-10-06.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __8103-11-04.pptx__ |                           |                | ..\extraction-testing-09102019\failed-ppts           |                          |
| __3112-11-12.pptx__ |                           |                | ..\extraction-testing-09202019\duplicate-errors-ppts |                          |
| __3112-12-02.pptx__ |                           |                | ..\extraction-testing-09202019\duplicate-errors-ppts |                          |











### Missing Data & Extra Blank Slide Lines in Extraction Sheet  

* [X] 1. Verify extraction-testing-09202019\duplicate-errors-ppts with Office 365
    
|File               | Verified in Office 365 (Yes/No) | Still Fails? (Yes/No) |
|-------------------|---------------------------------|-----------------------|
|3112-11-12.pptx    | Yes                             | Yes                   |
|                   | Yes                             | Yes                   |
|                   | Yes                             | Yes                   |
|3112-12-02.pptx    | Yes                             | Yes                   |
|                   | Yes                             | Yes                   |
|                   | Yes                             | Yes                   |
|3112-12-04.pptx    | Yes                             | Yes                   |
|                   | Yes                             | Yes                   |
|                   | Yes                             | Yes                   |
|8525-03-02.pptx    | Yes                             | Yes                   |
|                   | Yes                             | Yes                   |
|                   | Yes                             | Yes                   |
|8525-01-06.pptx    | Yes                             | Yes                   |
|                   | Yes                             | Yes                   |
|                   | Yes                             | Yes                   |
|8525-04-06.pptx    | Yes                             | Yes                   |

* [X] 2. Try moving the graphic and other files at the end of the ppt to the front of the ppt
* [X] 3. Retag no audio task slides

|File               | Fix Anything? (Yes/No) | Fixed What?  |
|-------------------|------------------------|--------------|
|3112-11-12.pptx    | Yes                    | Missing data, tagging: add No Audio: NA |
|                   | Yes                    | Empty slides at end of extraction sheet:<br>move graphics slide(s) to front info section |
|3112-12-02.pptx    | Yes                    | Empty slides at end of extraction sheet:<br>move graphics slide(s) to front info section |
|3112-12-04.pptx    | Yes                    | Empty slides at end of extraction sheet:<br>move graphics slide(s) to front info section |
|8525-03-02.pptx    | Yes                    | Empty slides at end of extraction sheet:<br>move graphics slide(s) to front info section |
|8525-01-06.pptx    | Yes                    | Missing data, tagging: add No Audio: NA |
|                   | Yes                    | Empty slides at end of extraction sheet:<br>move graphics slide(s) to front info section |
|8525-04-06.pptx    | Yes                    | Missing data, tagging: add No Audio: NA |

### No Vocab or Wrong Vocab

* [X] 1. Verify that previous vocab work was not lost in versioning
- * [X] Version v0.4.2-alpha - pull copy of code for "Beyond Compare" (extraction_pptx.py and extraction_out.py)
- * [X] Version v0.4.1-alpha - vocab code matches: Correct
- * [X] Version v0.4.0-alpha - vocab code matches: Correct

* [X] 2. Partial, wrong, or missing vocab - Find what issue is causing bug

|File               | Fixed? (Yes/No) | Fixed What?  |
|-------------------|-----------------|--------------|
| 3112-11-12.pptx   | Yes             | Missing vocab words - was using hyphen<br>instead ofdash replaced hyphen with dash<br>dash in words on<br>slide 22 replaced dash with colon |
| 3112-12-02.pptx   | Yes             | Missing vocab words - was using hyphen<br>instead ofdash replaced hyphen with dash |
| 3112-12-04.pptx   | Yes             | Missing vocab words - was using hyphen<br>instead ofdash replaced hyphen with dash |
| 8525-03-02.pptx   | Yes             | Wrong vocab words - dash in words on<br>slide 56 replaced dash with colon |
| 8525-01-06.pptx   | Yes             | Wrong vocab words - dash in words on<br>slide 24 replaced dash with colon |
| 8525-04-06.pptx   | N/A             | -            |


* [ ] 3. Partial, wrong, or missing vocab - Code fixes:
- * [ ] Use hypen or dash for the vocab word notes slides
- * [ ] Check if dash or hyphen content is vocab word(s)
- * [X] If content is not in a vocab notes slide, ignore
        as vocab and treat for whatever else it is instead

|File               | Fixed? (Yes/No) | Fixed What?  |
|-------------------|-----------------|--------------|
| 3112-11-12.pptx   |                 |              |
| 3112-12-02.pptx   |                 |              |
| 3112-12-04.pptx   |                 |              |
| 8525-03-02.pptx   | Yes             | Pulls only vocab from slide notes regardless of other non-vocab words with dashes |
|                   |                 |              |
| 8525-01-06.pptx   | Yes             | Pulls only vocab from slide notes regardless of other non-vocab words with dashes |
|                   |                 |              |
| 8525-04-06.pptx   | N/A             |              |
|                   |                 |              |