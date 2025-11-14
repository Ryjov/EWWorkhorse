This app is a microservice implementation of my other app E2W. It consists of 3 apps that communicate using RabbitMQ (I plan on replacing RMQ with gRPC communication in grpc branch): 
- EWWorkhorse - does actual business logic;
- EWeb - web app that allows user to download files for processing through a web interface;
- EWPFDesktop - desktop (WPF) app that allows user to download files for processing through a desktop application.

It allows user to transfer data from an Excel spreadsheet to a Word document. It does this by searching a Word document for special markers, that point to a cell in an Excel table, and replacing these markers with value from that cell.
Marker must follow this structure: <#(Worksheet number)#(Cell name)>. For example: <#3#B23> - this marker will point to cell B23 in 3rd worksheet of Excel table.

It uses an RPC system for sending a message from producer to consumer and sending an answer from consumer to producer.


EWWorkhorse
Console application that does the actual replacement.

Program.cs - creates and runs a Worker background service.

Worker.cs - creates a connection and declares a queue. Binds the queue to the exchange and listens for producers to send a message. Contains an event handler for a received message.
The message is the contents of either of 2 files send by producers in byte array form. Prefetch count is set to 2, since the producers will always send 2 files - Word document and Excel spreadsheet (probably not the best way to send two files, will look into it later).
The message with Word contents should always be send first and the one with the Excel contents is always second, so the app looks at the delivery tag number and checks if it can be divided by two to recognize the contents of which file have been sent (again, not the best solution, will have to check if there's a better way).
Event handler for a received message checks if both byte array variables for file contents have been filled. If they weren't the event handler finishes its work. 
If they were both filled it creates a new WordProcessingDocument and SpreadsheetDocument variables with received data and runs a ReplaceFile method from static Replacer.cs class that returns a new Word document with replaced values.
Then it transforms the resulting word document into a byte array to send a response to a producer.//

Replacer.cs - static class that contains methods to do the actual replacement.
ReplaceFile method is overloaded to either accept file contents as byte arrays or as WordprocessingDocument and SpreadsheetDocument.
To search for markers in text, a new Regex regular expression is created that will match needed text (<#\d+#[A-Z]+\d+>). Method finds regex matches by looping through all Text elements. Once it finds a match it looks for a sheet and cell number in a match with another Regex element. Using LINQ we Find needed sheet in Excel file and by going down its tree of objects we can find a needed cell value. If cell value is of numeric type we can simply retrieve its value, but if its of type string or bool then we need to do a data conversion. (for now lets assume that values in excel are either numeric or string - without boolean values) The cycle is then repeated for all detected regular expressions matches.


EWeb
Web application that allows user to download files for processing through a web interface.
