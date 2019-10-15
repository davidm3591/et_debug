# ET Debug Checklist - Current Release: v0.4.5-alpha

<hr>

## Working


### Bug: Codec Error Crash Files - Release: v0.4.5-alpha  (See end of doc for previous effort on this<sup>[(1)](#(1))</sup>)


Create package to post to PythonistaCafe, Stack Overflow
* [ ] 1. Desription of Issue
     * [ ] 1.1. Title of Problem
     * [ ] 1.2. Give a description of what the problem/issue is (as far as I can tell) in a text.txt file
     * [ ] 1.2. Give a description of what has been tried so far in a text.txt file
* [ ] 2. PowerPoint
     * [ ] 2.1. Rename 3112-09-02.pptx to 1234-56-78.pptx
     * [ ] 2.2. Clean out all proprietary content from 0123-45-67.pptx (Lorem-Ipsum)
     * [ ] 2.3. Strip out slides not needed to demo problem
* [ ] 3. Code
     * [ ] 3.1. Extract pertinent code that is broken - into a text.txt file
     * [ ] 3.2. Copy error message to text.txt file

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
   >1. Have found out that a lot of the statististical symbols are not available in Unicode as single characters. They have to be built combining character with a diacritical.
   >2. Some MS characters require something called UTF-16 surrogate pairs
   >3. __Biggie here__: Why Python seeing these objects as non-text characters??????
   
   #### What Should Happen:
   >Pull the content from a PowerPoint (Office 365) (using python-pptx ~~0.6.17~~ 0.6.18 (latest release)) regardless of whether it is text, a symbol, or an equation for output to Excel (Office 365) (using XlsxWriter 1.1.5) with Python 3.7?



________________
### Footnotes  


________________
### Archived Content  

<br/>

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