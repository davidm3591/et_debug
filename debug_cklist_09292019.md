# ET Debug Checklist

## Bug: 
_____________

# Completed
## Bug #X: Local Disk (C:\projects\extraction_tool_project\testing debug files\extraction-testing-09202019\duplicate-errors-ppts) - Release v0.4.2-alpha  

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


* [X] 3. Partial, wrong, or missing vocab - Code fixes:
- * [X] Use hypen or dash for the vocab word notes slides
- * [X] Check if dash or hyphen content is vocab word(s)
- * [X] If content is not in a vocab notes slide, ignore
        as vocab and treat for whatever else it is instead

|File               | Fixed? (Yes/No) | Fixed What?  |
|-------------------|-----------------|--------------|
| 3112-11-12.pptx   | Yes             |              |
| 3112-12-02.pptx   | Yes             |              |
| 3112-12-04.pptx   | Yes             |              |
| 8525-03-02.pptx   | Yes             | Pulls only vocab from slide notes regardless of other non-vocab words with dashes |
|                   |                 |              |
| 8525-01-06.pptx   | Yes             | Pulls only vocab from slide notes regardless of other non-vocab words with dashes |
|                   |                 |              |
| 8525-04-06.pptx   | N/A             |              |
|                   |                 |              |