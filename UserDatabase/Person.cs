using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UserDatabase
{
    class Person
    {
        public static string GetFirstName()
        {
            // gets first name
            // variables to hold user input
            string firstName = "";
            // prompt user to input
            Console.Write("Enter first name: ");
            firstName = Console.ReadLine();
            return firstName;
        }
        public static string GetLastName()
        {
            // gets first name
            // variables to hold user input
            string lastName = "";
            // prompt user to input
            Console.Write("Enter last name: ");
            lastName = Console.ReadLine();
            return lastName;
        }
        public static string GetPrompt()
        {
            // asks the user if they want to add more names
            // variables to hold user input
            string prompt = "";
            // prompt user to input
            Console.Write("Do you want to add another name? Y/N ");
            prompt = Console.ReadLine();
            return prompt;
        }
    }
}
