Python and main modules
Python 3.7.1
python - pptx 0.6.17
XlsxWriter 1.1.5

Info:
    The alpha symbol - Unicode Hex: 1D6FC
        (Font from insert > symbol says normal font - Cambria Math is
            selected font)
    The p (as symbol) - Unicode Hex: 1D45D
        (Font from insert > symbol says normal font - Cambria Math is
            selected font)
    The p-hat symbol - Unicode Hex: 1D45D for the p, but the hat is a
        diacritical that is combined with the p
        (Font from insert > symbol says normal font - Cambria Math is
            selected font)



Exception in Tkinter callback
Traceback(most recent call last):
    File "c:\users\david\anaconda3\Lib\tkinter\__init__.py", line 1705, in __call__
        return self.func(*args)
    File ".\extract_build\extraction_pptx.py", line 205, in get_file
        read_slide(filename_path)
    File ".\extract_build\extraction_pptx.py", line 356, in read_slide
        for line in slide.notes_slide.notes_text_frame.text.split("\n"):
AttributeError: 'NoneType' object has no attribute 'text'

Notes:
Found out that a lot of the statististical symbols are not available in
    Unicode as single characters. They have to be built combining a
    character with a diacritical.
Found out that some MS characters require something called UTF - 16 surrogate pairs


Python
filename_path = glob.glob("C:\\edge_tool_data\\input\\*.pptx")
.
.
.
tags = re.compile(r"^#[a-zA-Z0-9\-]+")
# The dash separator is an endash to identify vocab words (not allowed
# anywhere else in the pptx slide notes)
vocab = re.compile(r"(([\w,-]+ ){1,})–(.+)")
# Audio (au) identified by appended text: 'hint', 'exit', etc.
.
.
.
if slide.has_notes_slide:
    # Get notes from PPTX slides
    # https://python-pptx.readthedocs.io/en/latest/user/notes.html
.
.
.
    # Run try except block for NoneType error
    try:
        for line in slide.notes_slide.notes_text_frame.text.split("\n"):

            # Match for Vocab
            v = vocab.match(line.strip())
            if v is not None:
                if chk_vocab_notes in slide.notes_slide.notes_text_frame.text:
                    vocab_results[v[1]] = v[3]

            # Tags preceded by hash tag (#)
            t = tags.match(line)
            if t is not None:
                tags_result.append(t[0])

            # Match for Audio
            au = audio.match(line.strip())
            if au is not None:
                audio_results[au[1]] = au[2]

    except Exception as e:
        file_warning_list.append(slide_num)
        text_error_flag = True
        # print(e)
