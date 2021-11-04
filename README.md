# csv-reader

The newline symbol is CR/LF(0x0d0a) or CR(0x0d) only in the csv file.

and the contents of some fields contain CR(0x0d), that cells must be enclosed double quote symbol.

thus, CR(0x0d) symbol is newline only or field contents or newline with LF.



1st : the OpenFileDialog look for the products master csv file that file contain product barcode.

2nd : the OpenFileDialog look for the packing list file that file ask to insert product barcode.

3th : the OpenFileDialog look for the saving file name that file save the packing list with barcode.
