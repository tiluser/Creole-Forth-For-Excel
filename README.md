Creole Forth for Excel
----------------------

Intro
-----

This is a scripting language built on top of VBA for Excel based on the Forth language.
The design and architecture of the language is similar to a previous language that worked
as a Delphi/Lazarus component. 

How to run
----------

1. Open up the CFExcel1.xls workbook. If you do not have Excel, there's a standalone version
   in the same folder CFExcel1.exe that can be used instead.

2. Make sure macros are enabled.

3. Type the shortcut CTRL-Shift-C. Two sheets will appear that weren't there before : one with
   a GlobalDS label and the other with a Dictionary label. 

4. Go to the GlobalDS sheet.

5. In the cell next to the label 'Input Area', enter two words: HELLO WORLD.

6. Hit the submit button.

7. Two dialog boxes will pop up : the first says "Hello" and the second says "World".

Compiling programs
------------------

Like other Forths, Creole has a colon compiler. Below are two example programs:

: SQR DUP Nx ;

: TESTBU BEGIN 1 N+ DUP 10 GT UNTIL ;

If you paste these programs into the Input Area and hit submit, you will see their entry into 
the dictionary. To execute, put the following in the input area

2 SQR

The above will place 2 on the stack, duplicate itself, and then multiply the two operands, 
leaving 4 on the stack.


TOGGLESU 1 TESTBU TOGGLESU

The TOGGLESU command toggles screen updating (turns it off if it's on and vice versa)
and should be used whenever there's a word with a looping construct (it's very slow otherwise).
1 TESTBU puts the number 1 on the stack, and adds it to itself until it goes over the limit
specified in the definition (10).   

Some questions 
--------------

Q: What is Forth? 

Forth is a stack-oriented language primarily known for programming hardware and embedded systems. It's 
distinguished by its simple postfix syntax, passing of parameters via a stack, and miniscule size and 
use of resources.    

A: How does Creole work?

1. A set of space-delimited values (words) is placed in an input-cell.
2. One by one, they're looked up in a dictionary.
3. If found, they're executed.
4. If not, they're pushed onto the data stack. 

Q: What version of Excel should I be using?

A: It should be backward compatibile to version 2003.

Q: What if I don't have Excel?

A: There's a version of available compiled to an executable using XLtoEXE (a free tool) in the same
folder as the spreadsheet.  

Q: Does it work with OpenOffice/LibreOffice?

A: No, at least not yet. It seems likely that a separate version would have to
be developed to work with them.  


Q: The programs and examples given seem awfully simple. Can't it do more?

A: Perhaps, but it's already demonstrated one of the ways you can improve and extend the Creole Forth environment.

There are several ways to extend the language but we'll mention two for starters:

1. By creating additional "primitives". A primitive consists of a VBA class method that is defined in one of the class modules
   attached to the spreadsheet. It must then be introduced into the Dictionary page via the BuildPrimitive method. 

2. By adding high-level definitions. This is generally done through the colon compiler. Its job is to compile assemblages
   of previously defined "words" or definitions. These behave as extensions to the language and are functionally indistinguishable 
   from lower level (primitive definitions). 

List of shortcuts and commands
------------------------------
1. Ctrl-Shift-C : Sets up default Forth Bundle, which consists of a GlobalDS page and a Dictionary page.
   Most of the time, you're only going to need one bundle. 

2. Ctrl-Shift-R : Removes default Forth Bundle.

3. Push button : Put a value in the Scratch cell,  and press the Push button. The value will be placed on
   top of the DataStack.

4. Submit button : Submits code in the InputAea.

5. Clean up stacks button: Push it and it will empty the DataStack, ReturnStack, and ParsedInput.

Things to watch out for, especially for experienced Forth programmers
---------------------------------------------------------------------

1. Dates - Excel tries to convert anything it can into a date and turning off this feature is very difficult. I would avoid putting anything into the input or stacks that can be converted into one. 

2. Check the dictionary to see if any of the common primitives have been renamed. 

3. N+ replaces + because Excel will interpret a cell with a leading + as a formula. Same is true for N-.
   

4.  &#42; has become Nx due to trouble with the lookup process with *.

5. Boolean comparison operators such as < , <=, >, =, etc are now LT, LE, GT, EQ, etc.

6. Be careful about deleting any named ranges on the worksheets - they're the key to everything working. You should use the push button to push items on the stack, don't just do it manually. 

7. If something gets screwed up manually, the Ctrl-Shift-R/Ctrl-Shift-C will bring everything to its previous state. I generally prefer it now to using FORGET, although that's available.



Changes by date
---------------
12/17/2016
1. Added the AppSpec/AppSpecMain modules to ease building of new primitives for the user and made APPSPEC the current default vocabulary. 

2. Added "How to build primitives in creole forth for excel" PowerPoint presentation.

3. Added a TutorialCode class module. It has examples from the presentation above
   to paste into the AppSpec module and isn't meant to be used directly.

4. Added list compiler { } which allows multiple arguments to be pushed onto the stack
   and occupy one cell. 

5. Set up the ExecForthWords sub to allow execution of CFE words in a string passed as a parameter. 

12/07/2016

1. Added Forth Day 2016 PowerPoint presentation. 

08/29/2016

1. Updated README.md file with "Things to watch out for" and "Changes by date".

2. Made some minor formatting changes to README.md. 

8/28/2016

1. For those who don't have Excel, there's a version compiled into and exe using XLToEXE, a free tool. It's in the repository as a zip file
   and just has to be unzipped an run (no instalation required). It appears to work the same as the regular spreadsheet and you can even see 
   the source code of the project. 

2. There's a README.md file now. It's intended to replace the quick_intro_cs.txt file. 

08/26/2016

1. Toggling of screen updating in rebuild (is now many times faster)

2. Addition of RunCommand sub. It allows features such as compiling and running code from an arbitrarily located cell.

3. CONSTANT and VARIABLE are now a default part of the dictionary (defined as high-level definitions).

4. Fixed DoDoes so DOES> works correctly when called from inside a colon definition. Before is only worked right when
   the interpreter handled it.

5. DEPTH fixed.

6. Empty stacks button. Removes junk from the places they're most likely to accumulate and cause trouble. Specifically
   the Data Stack, the Return Stack, and the Parsed Input.

7. Added the TOGGLESU primitive. Allows alternatively turning off screen updates, and then turning back on. Great to use
   in looping statements, which are slow otherwise.

8. TOGGLESHOW primitive. This allows alternately hiding and showing the code sheets. Right now the GlobalDS and Dictionary
   pages are there, but others could be added at the discretion of the developer. 

9. Changed NOW to pop up a message box instead of pushing a value onto the stack. I didn't like its tendency to change formatting everywhere it 
   touched. 

10. Set limit of ForthBundleCount to 4 - it's now a property of the ForthBundleParamSet.

08/07/2016
Initial commit. 
