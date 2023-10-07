# picturesVBA

EXCEL VBA pictures insert

**The problem is how to lookup pictures to an excel sheet by item codes.**

We are making an order list for clients. There is a pack of photos for each article, each photo named itemcode.jpg and we need to look up a photo for each itemcode.

There are several way to do it, based on where photos are stored.

 
  1.a) We could upload photos to a server and use functoin =IMAGE(https://links.example/itemcode.jpg). For the method we need certain link in HTML format, **you can not use** links from sharepoint, google drive or one note.     
     

  Pros of the method - you don't need upload pictures into the file, images are not compressed and the quality depends on cells row and height;
  
  Cons -  we need server to upload pictures and be sure that client have access to the server; clients version of microsoft office could be old and the IMAGE function could be unavailable;
  
  For the method just use a list of itemcodes and links to reffered image. Look up from the table a link into a nested image function like
     
       =IMAGE(XLOOKUP(itemcode, range_of_articles, range_of_links))

     
  1.b) To avoid problem with microsoft version you can use Google sheet with the same =IMAGE() function, but you still need server to upload pictures.



2) We could lookup pictures from the local disk, with path like D:/pictures/itemcode.jpg
     Pros - pictures will be stored into a file, a client doesn't need access to the server;
     Cons - the file size is growing, to keep it within 20 mb (limit for Outlook email attachment) you need to compress pictures;.

In my case, I couldn't use any server to upload pictures to, so I use the second way.

    
     
