<a name="(top)" style="color: black; text-decoration: none;">

(Top)

</a> 

# ET Debug Checklist - Current Release: v0.4.5-alpha


<hr>

## Working


### Bug: Codec Error Crash Files - Release: v0.4.5-alpha  (See end of doc for previous effort on this<sup>[(1)](#(1))</sup>)

#### From PythonistaCafe post<sup>[(2)](#(2))</sup> - it appears that adding either equations or Cambria Math changes object from notes_text_frame to something else - i.e., no longer inherits text_frame property?

#### (NOTE: pick up with 'Extract all text from slides in presentation')

* [ ] 1. Research & Testing - build resource list with links
     * [ ] 1.1. Resource link list
        * [ ] 1.1.1. [python-pptx Full Documentation - Release v0.6.18](https://python-pptx.readthedocs.io/en/latest/index.html#api)
        * [X] 1.1.2. [python-pptx - Full Documentation - Release v0.6.18 - PDF](https://buildmedia.readthedocs.org/media/pdf/python-pptx/latest/python-pptx.pdf)
        * [X] 1.1.3. [Getting Started - Building a PPT](https://python-pptx.readthedocs.io/en/latest/user/quickstart.html)
        * [ ] 1.1.4. [MSO_AUTO_SHAPE_TYPE](https://python-pptx.readthedocs.io/en/latest/api/enum/MsoAutoShapeType.html#msoautoshapetype)
        * [ ] 1.1.5. [Working with placeholders - Identify and Characterize a placeholder](https://python-pptx.readthedocs.io/en/latest/user/placeholders-using.html)
        * [ ] 1.1.6. [Working with Notes Slides](https://python-pptx.readthedocs.io/en/latest/user/notes.html)
        * [ ] 1.1.7. [Font typeface](https://python-pptx.readthedocs.io/en/latest/dev/analysis/txt-font-typeface.html)
        * [ ] 1.1.8. [Python - Unicode HOWTO](https://docs.python.org/3/howto/unicode.html)
     * [ ] 1.2. Set up test files, stripped down (base file: pptx file that was submitted to PythonistaCaf)
        * [X] 1.2.1. Stripped out with symbol, equation, Cambria Math (CM) (slides 4, 5, 7 - symbols/equations)<br>
                     0123-45-67_stripped_w_symbol_equation .pptx
        * [X] 1.2.2. Replace slide completely - no symbol, CM, equation (slides 4, 5, 7 - from 3316-09-04_passes.pptx)<br>
                     0123-45-67_replace_slide_symbol_equation.pptx
        * [X] 1.2.3. Replace slide completely - Just add CM, to a number or something (slide 4 - x y z in CM)<br>
                     0123-45-67_replace_slide_just_add_cambria.pptx
        * [ ] 1.2.4. Replace slide completely - Just add a plain text symbol (slides 4, 5, 7 - add different symbol/pg)<br>
                     0123-45-67_replace_slide_just_add_symbol.pptx
        * [X] 1.2.5. Clean file - passess with vocab<br>
                     3316-09-04_passes.pptx
     * [ ] 1.3. Set up extraction_pptx.py to read __every iteration__ and output the
        * [ ] 1.3.1. slide.notes_slide.notes_text_frame
        * [ ] 1.3.2. placeholder.placeholder_format.type
        * [ ] 1.3.3. shape in slide.notes_slide.shapes
        * [ ] 1.3.4. shape.text
     * [ ] 1.4 . Subsitute and test
        * [ ] 1.4.1. new vocab slide w/notes that has not symbols or equations
        * [ ] 1.4.2. change something in new slide to Cambria Math font
        * [ ] 1.4.3. if font change didn't break it, insert a symbol with arial font
* [ ] 2. Fix - Implement fix for symbol and equation read bug
    * [ ] 2.1 
_________________

#### Table of files

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


### Problem-Error Information
   #### To Reproduce (specific software and modules):
   >Use (a) Python 3.7, (b) python-pptx 0.6.17, and (c) XlsxWriter 1.1.5 with split() on new line (`for line in slide.notes_slide.notes_text_frame.text.split("\n"):`), to grab the notes slide content to output to an Excel spread sheet word-for-word (including symbols and equations).
   >
#### Current Error:
   >When a symbol (like p-hat or x-bar), equation or formula (font: Cambria Math) is encountered in a notes slide, it throws the following error:
   <pre>
    Exception in Tkinter callback
    Traceback (most recent call last):
      File "c:\users\david\anaconda3\Lib\tkinter\__init__.py", line 1705, in __call__
        return self.func(*args)
      File ".\extract_build\extraction_pptx.py", line 205, in get_file
        read_slide(filename_path)
      File ".\extract_build\extraction_pptx.py", line 356, in read_slide
        for line in slide.notes_slide.notes_text_frame.text.split("\n"):
    AttributeError: 'NoneType' object has no attribute 'text'
   </pre>

#### Notes:
   >1. __*Biggie here*__: Why is Python seeing these objects as non-text characters??????
   >2. Found out that a lot of the statististical symbols are not available in Unicode as single characters. They have to be built combining a character with a diacritical.
   >3. Found out that some MS characters require something called UTF-16 surrogate pairs
   
   #### What Should Happen:
   >Pull the content from a PowerPoint (Office 365) (using python-pptx ~~0.6.17~~ 0.6.18 (latest release)) regardless of whether it is text, a symbol, or an equation for output to Excel (Office 365) (using XlsxWriter 1.1.5) with Python 3.7?



________________
### Footnotes  


________________
### Archived Content  

<br/>

[(top)](#(top))

<a name="(1)" style="color: black; text-decoration: none;">

#### (1) Previous Effort: Bug: Codec Error Crash Files - Release: v0.4.3-alpha

</a> 

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

* [X] 1. State the problem (see Problem-Error Statement below)

### Problem-Error
   #### To Reproduce:
   >Use (a) Python 3.7, (b) python-pptx 0.6.17, and (c) XlsxWriter 1.1.5 with split() on new line (`for line in slide.notes_slide.notes_text_frame.text.split("\n"):`), to grab the notes slide content to output to an Excel spread sheet word-for-word (including symbols and equations).
   >
#### Cuurrent Error:
   >When a symbol (like p-hat or x-bar), equation or formula (font: Cambria Math)) is encountered in a notes slide, it throws the following error:
   <pre>
    Exception in Tkinter callback
    Traceback (most recent call last):
      File "c:\users\david\anaconda3\Lib\tkinter\__init__.py", line 1705, in __call__
        return self.func(*args)
      File ".\extract_build\extraction_pptx.py", line 205, in get_file
        read_slide(filename_path)
      File ".\extract_build\extraction_pptx.py", line 356, in read_slide
        for line in slide.notes_slide.notes_text_frame.text.split("\n"):
    AttributeError: 'NoneType' object has no attribute 'text'
   </pre>
   
   #### What Should Happen:
   >Pull the content from a PowerPoint (Office 365) (using python-pptx ~~0.6.17~~ 0.6.18 (latest release)) regardless of whether it is text, a symbol, or an equation for output to Excel (Office 365) (using XlsxWriter 1.1.5) with Python 3.7?

* [X] 2. Modify Try/Except block to output error traceback msg, slide number
* [X] 3. Identify offending equation or symbol
#### Example file (x-bar and p-hat): __3112-11-12.pptx__

  >frame_content_warning.txt content:
  >
  >SLIDE CONTENT ERROR WARNING LIST
  >
  >The following slides have text formatting problems. Please review the Extraction Sheet for content errors.
  >
  >Slide: 8
  >
  >Possible format errors may include equation elements or non-standard symbols. Please compare problem slides to the Extraction Sheet for content errors. If data in the Extraction Sheet is incorrect or missing, the slide formatting errors will need to addressed.
* [ ] 4. Research solutions
  - * [X] Create new virtualenv with newest release of python-pptx (v0.6.18)
      - ~~Try extraction with updated python~~ NOT THE ISSUE
  - * [X] Search for info on same issue(s)
      - Found that some Cambria Math symbols are outside of UTF-8 (need 16 bit UTF-16 surrogate pairs)

      >Source: https://stackoverflow.com/questions/5831571/textout-and-the-cambria-math-font
      >
      >The Cambria Math font has UNICODE characters beyond 0xFFFF. You can see them in a Word document, just by inserting a Symbol and selecting the Cambria Math font. By the way, the Windows Character Map does not show these characters. My question is : how to exhibit those UNICODE characters in a Windows app using TextOut() ?
      >
      >To display these supplementary code points you need to use UTF-16 surrogate pairs.<br>
      >A surrogate pair is a way of representing single code points beyond 0xFFFF as two wide characters. You simply pass a surrogate pair to TextOut() and it will be displayed.

  - * [ ] Soooo, how to solve(???) (textout() is a windows method)
      - Use UTF-16 surrogates in Python 3.x - is there a module or maybe a 'method/function'? research? post?
      - Possible to parse and treat as image? (Is in-line-conversion possible? Python does not even see it as a text-object...)
  - * [ ] Use solution in the exception block?


<br/>

________________

[(top)](#(top))

<a name="(2)" style="color: black; text-decoration: none;">

#### (2) PythonistaCafe Post Effort: Bug: Codec Error Crash Files - Release: v0.4.5-alpha

</a> 

#### Complete - answer on slide.notes_slide.notes_text_frame - may not be codec issue

#### Create package to post to PythonistaCafe, Stack Overflow
* [X] 1. Desription of Issue
     * [X] 1.1. Title of Problem
     * [X] 1.2. Give a description of what the problem/issue is (as far as I can tell) in a text.txtc file
     * [X] 1.3. Summarize the issues with the content that is crashing
     * [X] 1.4. Screen shots of pptx with highlighting and text explaining wht it is
* [X] 2. PowerPoint
     * [X] 2.1. Rename 3112-09-02.pptx to 0123-45-67.pptx
     * [X] 2.2. Clean out all proprietary content from 0123-45-67.pptx (Lorem-Ipsum)
     * [X] 2.3. Strip out slides not needed to demo problem
* [X] 3. Code
     * [X] 3.1. Extract pertinent code that is broken - into a file
     * [X] 3.2. Copy error message to the file
     * [X] 3.3. Add Python version and major modules to file
     * [X] 3.1. Extract pertinent code that is broken - into a text.txt file
* [X] 4. Review posting-to videos
* [X] 5. Review/CE post content and post