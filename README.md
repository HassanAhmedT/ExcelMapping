# ExcelMapping

A .Net Core based library to map excel record into generic object.

## Usage

```c#
// create your class model

public class Book
    {
        public int BookID { get; set; }
        public string BookName { get; set; }
    }

// create the validator that implement the IValidationDataProperty if needed to validate
// for each row in the excel and map only that return true.

public class EmptyFieldValidation : IValidationDataProperty
    {
        public bool CheckPropertyValue(string value)
        {
            return value.IsNotEmpty();
        }
    }

// create the parser that implement IParsingDataProperty if needed to 
// parse the value in the excel sheet to the desired format of data.

public class ParsingIntValues : IParsingDataProperty
    {
        public object ParsePropertyValue(string value)
        {
            return int.Parse(value);
        }
    }

// create the mapping class that inherit from the main class 
// map each column with its equivalent property and add validator and data 
// parsing if needed

public class MappingBookExcel : ExcelMapper<Book>
    {
        public MappingBookExcel()
        {
            // mapping the column text as 'bookid'
            // and 
            // the property to store the value in as 'a => a.BookID'
            MappingColumn("bookid", a => a.BookID, new ParsingIntValues(), new EmptyFieldValidation());
            MappingColumn("book name", a => a.BookName, new EmptyFieldValidation());
        }
    }

// finally create an object from that last class and execute the .Map() function
// and pass the ExcelWorkSheet to it and it return List of Books.
new MappingBookExcel().Map(workSheet);
