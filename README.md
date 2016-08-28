Creole Forth for Excel

Intro

This is a scripting language built on top of VBA for Excel based on the Forth language.
The design and architecture of the language is similar to a previous language that worked
as a Delphi/Lazarus component. 

How to run
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
Ctrl-Shift-C : Sets up default Forth Bundle, which consists of a GlobalDS page and a Dictionary page.
Most of the time, you're only going to need one bundle. 
Ctrl-Shift-R : Removes default Forth Bundle.
Push button : Put a value in the Scratch cell,  and press the Push button. The value will be placed on
top of the DataStack.
Submit button : Submits code in the InputAea.
Clean up stacks button: Push it and it will empty the DataStack, ReturnStack, and ParsedInput.
