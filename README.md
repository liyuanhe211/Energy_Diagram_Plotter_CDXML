# Energy Diagram Plotter CDXML
A tool to create pixel-accurate energy diagrams as ChemDraw object.

In the literatures, I see lots of people use the ChemDraw+mouse+eye method to draw energy diagrams. It is tedious and sometimes extremely inaccurate. Even the order of energy could be accidentally inverted. Which could be very misleading. The left figure below is the line graph drawn by ChemDraw in the literature, and the right is the to-scale version (ignore the unit change from kcal/mol to kJ/mol). The original diagram is a mess. Such diagrams are even worse than only giving numbers as a table.

![image](https://user-images.githubusercontent.com/18537705/160621653-d8dd5409-9351-4fda-81aa-b602a1f8afd2.png)

There are lots tools to generate energy diagrams. For example DataGraph or [mechaSVG](https://github.com/ricalmang/mechaSVG). They are OK for basic graphs, but they usually genreates an un-editable figure with very limited customization options. Also one cannot drag the numbers and tags, trubblesome for complex energy diagrams.

The program here can automatically generate ChemDraw files with line charts based on simple Excel input. After that, you can freely modify the ChemDraw elements. And it's very convient to add chemical structure to it.

The input file format used by the program is as follows, and the left-right comparison should be easy to understand:
Any edits you want can then be done in ChemDraw:

The supported colors are expressed as follows (actually the colors supported by Matplotlib http://matplotlib.org/api/colors_api.html)

The interface and the meaning of each option are described in the following figures:
The font of the figure is fixed, and the size and aspect ratio of the figure are determined by the window size

Note that some of the following options are only reflected in the ChemDraw file, not in the preview view.

The last "Use temp file" option is to decide whether the program creates a file in its own Temp folder (the user needs to save it after viewing it), or directly generates a cdx file in the directory where Excel is located.

Other options not mentioned should be self-explanatory.


Precautions

A known bug: The Tag in Excel must be in text mode, so if the tag is simply a number (such as calling a state "3"), a type error will be reported. It can be temporarily solved by adding spaces around the pure number label, or manually changing its cell properties to text mode.


The temporary files of the program are stored in the program directory, so the program must have read and write permissions to this directory, so it is not recommended to put the program in a path with permission restrictions such as Program Files, otherwise please give it administrator permissions.


The software has only been tested under Win 8.1. Win 10 can try it, it should be no problem (if it prompts that DLL is missing, apply the last patch of the post). 32-bit operating systems are not supported.


The preview interface is not completely correct. If there is unexpected behavior, export the ChemDraw file first to see if it is normal.


The program's "Avoid text overlap" function is currently greedy, and may produce unreasonable results in particularly crowded situations. Just fine-tune it yourself (tip, hold down Shift and drag the text in ChemDraw to ensure it is horizontal or vertical). move).

Questions and Feedback (Bug Report)

This program is self-use software and may have other bugs. If you have any questions, please see the Example in the program directory first.
After that, you can give feedback directly in this post, and upload your own Excel input file, the last display in the CMD window, and a screenshot of the operation interface. If the output file is generated but there is a problem with the output file, please upload the output file as well. Otherwise, it is not easy to repeat the question based on the description alone, and no reply will be given.
If there is a "flashback", please run the program through the "Draw_Energy_Diagram_XML_Debug.bat" file in the directory, so as to keep the error message before the program exits. (The old version does not have this file, you can download the following .bat and put it in the directory where the program is located to run)


![image](https://user-images.githubusercontent.com/18537705/160620003-5657e605-e95c-495b-aeae-b43006e78b6b.png)

![image](https://user-images.githubusercontent.com/18537705/160621422-05274905-5b1e-43b7-8cde-1ae80577d795.png)


![image](https://user-images.githubusercontent.com/18537705/160621753-f0bad86a-5238-4f72-b7b0-c4554e0c0e0d.png)

![image](https://user-images.githubusercontent.com/18537705/160618917-1544b242-6f63-45de-96e9-3892b549d445.png)


![image](https://user-images.githubusercontent.com/18537705/160621816-316c6304-1598-4a9c-83dc-666a071d2cb8.png)

![image](https://user-images.githubusercontent.com/18537705/160621848-128372f6-c51b-4674-bb1e-4a50adf370cb.png)

![image](https://user-images.githubusercontent.com/18537705/160621896-bbc03f1c-df92-43d2-8d60-0869156daec3.png)

![image](https://user-images.githubusercontent.com/18537705/160621889-39ee11a1-b997-48bd-b7ae-99887a75da99.png)

![image](https://user-images.githubusercontent.com/18537705/160621922-7f7f0526-fcb3-4df6-bbd2-3db4f1c1736c.png)

![image](https://user-images.githubusercontent.com/18537705/160621958-1d139beb-878b-43b8-b882-4dbf56439ff0.png)
