using IronXL;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ReadDataFromCsvFile
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Searching Program!");
            Console.WriteLine("Enter X To Exit!");
            bool IsExit = false;        //Property to stop the execution of program bydefault set to false

            while (IsExit == false)
            {
                Console.Write("\nEnter Key To Search Data : ");     
                var search = Console.ReadLine();            //Get User Input
                if (search.ToLower()!="x")              //If User Input Is Not X or x
                {
                    var PersonData = GetPersonData(search);     //call the function of searching and passing the search value as parameter
                    if(PersonData!= null)                       //if search matches the data
                    {
                        Console.WriteLine("\nPerson ID : " + PersonData.ID);                    //Show Person ID
                        Console.WriteLine("Person First Name : " + PersonData.FirstName);       //Show Person First Name
                        Console.WriteLine("Person Last Name : " + PersonData.LastName);         //Show Person Last Name
                        Console.WriteLine("Person Date Of Birth : " + PersonData.DOB);          //Show Person Date Of Birth
                    }
                    else        //If Data Not Found With Matching Key
                    {
                        Console.WriteLine("No Data Found With This Key");
                    }
                }
                else    //If User Type X to stop the program execution
                {
                    IsExit = true;
                }
            }

            Console.WriteLine("Press Any Key To Exit");
        }

        public static List<Person> ReadCSVData(string csvFileName)
        {
            //var csvFilereader = new List<string>();
            var csvFilereader = ReadExcel(csvFileName);     //call the read function and get list of strings for the rows of worksheet

            List<Person> Person = new List<Person>();   //initialize list of class Person class
            foreach (var i in csvFilereader)
            {
                Person person = new Person();   //initialize object for class Person
                var mydata = i.Split(',');      //Split the values by comma

                person.ID = mydata[0];          //Add First Value Of to Person Class ID Property
                person.FirstName = mydata[1];   //Add Second Value Of to Person Class FirstName Property
                person.LastName = mydata[2];    //Add Third Value Of to Person Class LastName Property
                person.DOB = mydata[3];         //Add Fourth Value Of to Person Class DOB Property

                Person.Add(person);             //Add The Object In The List Of Person
            }
            return Person;      //return the entire list containing the objects of class Person
        }
        public static Person GetPersonData(string searchBy)
        {
            var Data = ReadCSVData("../../../files/Data.csv");  //Call The Function Containing the list of objects for class persons
            return Data.Where(x => x.FirstName == searchBy || x.LastName == searchBy || x.DOB == searchBy || x.ID == searchBy).FirstOrDefault();//Searching by the user input key
        }
        private static List<string> ReadExcel(string fileName)
        {
            WorkBook workbook = WorkBook.Load(fileName);
            // Work with a single WorkSheet.
            //you can pass static sheet name like Sheet1 to get that sheet
            //WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
            //You can also use workbook.DefaultWorkSheet to get default in case you want to get first sheet only
            WorkSheet sheet = workbook.DefaultWorkSheet;

            List<string> newdata = sheet.Rows.Cast<object>().Select(o => o.ToString()).ToList();     //Convert the worksheet rows to List of strings
            return newdata;     //return the List of strings
        }

        public class Person
        {
            public string ID { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string DOB { get; set; }
        }
    }
}
