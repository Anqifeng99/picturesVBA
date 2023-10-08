# picturesVBA
EXCEL VBA pictures insert
To insert pictures into excel file you can use two methods

1) worksheets("sheet1").pictures.insert(FileName)

2) worksheets("sheet1").shapes.adppicture "FileName", LinkToFile, SaveWithDocument, Left, Top, Width, height

If you record macro and insert picture from the insert tab in excel, you will get the first method in the recorded code. But the picture will not be saved into the file, it will be just a link to the pcture on the disk. 
If you try share the excel sheet with such links, a person, you sent file to, couldn't see pictures. To save picture in the excel sheet we should use shape.addpicture() method


shapes.addpicture method is a little bit tricky, because you can't just choose shape "picture" from the shape list in excel. The method has 7 arguments:

  1) FileName with a path to a file in double quotation marks, for example "D:/pictures/10001.png";
  2) LinkToFile, true if you want just keep links to file or False if you want to get pictures as object in the excel workbook. In our case I set False;
  3) SaveWithDocument, true if you want to keep pictures into a woorkbook.
  4) left - left coordinate. You can use a number, or reference to a cell, like activecell.left or cells(1,1).left
  5) top - the top coordinate.As for left, you can use a number or a reference, like cells(1,1).top
  6) width - width of a picture in pix. Use -1 to keep original size
  7) height - height of a picture in pix. Use -1 to keep original size

  Important notes:
  1) you dont need brackets to the method. Just write a file name in quotation mark after shapes.adppicture 
  1) path to files, excel workbook and pictures, MUST BE ONLY IN ENGLISH LETTERS. 
      There could be an error, if the path contains cyrillic, arab, hindi letters, latin letters with diacritic, chinese characters etc.
  2) if your pictures has different heights and width, I recomend insert them into the file with a original sizes, (-1 for width and -1 for heigth), and then select all pictures and set a necessary hight to all. It helps to keep ratio of pics. 
  3) The file with pics will be heavier, then with links, so you should find ballance between weight of the file and pics quality. If you need to send the file by outlook, it must be not more than 20mb.  
