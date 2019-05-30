# ApachePOI-POC
Proof of Concept on how to use Apache POI java library to read/write Excel files. Shuffling existing records and generating new records satisfying monthly target count sheets, and feeding them into a dummy MS SQL database using a JDBC connection.

This project is a special use case customised for the databases of one of my clients to create a dummy for their 3 years' lost records using only 3 months of left records and assumes a certain backend DB schema and layout of input excel workbooks.

The final version is in https://github.com/glassrose/ApachePOI-POC/tree/master/src/WeeklyOPD
