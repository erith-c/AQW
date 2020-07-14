#define CONSOLE
#define FILE
#define EXCEL

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[Author("Erith", version = "0.0.1")]

// Global Attributes
[AttributeUsage(AttributeTargets.All,
                AllowMultiple = true)]
public class Author : Attribute
{
    public string name { get; }
    public string version;

    public Author(string name)
    {
        this.name = name;

        // Default value
        version = "1.0.0";
    }
}

// Attributes local to AngelLight Software
namespace ALS
{
    [AttributeUsage(AttributeTargets.All)]
    class WARDENAttribute : Attribute
    {
        public string version { get; set; }
        string author { get; set; }

        public WARDENAttribute(string version, string author = "Erith")
        {
            this.version = version;
            this.author = author;
        }
    }
    
    // Attributes local to Project AQUA
    namespace AQUA
    {
        [AttributeUsage(AttributeTargets.All)]
        public class Attributes : Attribute
        {

        }
        public class Target : Attributes
        {

        }
    }
}
