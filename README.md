editbot
=======

Let robots do the copyediting.

(Eventually, this project will include Python scripts and Microsoft Word 
macros to automate the most tedious copyediting tasks. For now, I'm collecting 
examples of patterns to match and replace and working on the Word macro. The 
patterns live in `patterns.py`. The Word macro lives in `editbot.dotm`.
Feel free to contribute.)

## Microsoft Word

### Running the macro

* Navigate to `Developer > Add-Ins > Add...`
* Add `editbot.dotm`
* Open the document you want to edit
* Navigate to `Tools > Macro > Macros...`
* Under `Macros in:`, select `editbot.dotm (global template)`
* Run the macro named `editbot`

### Editing the macro

* In Word, do `File > Open... > /path/to/editbot.dotm`
* `Developer > Editor`
* The macro named `editbot` is under `Module1`
