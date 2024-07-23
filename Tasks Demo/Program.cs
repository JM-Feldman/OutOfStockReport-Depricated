using System.Threading.Tasks;
namespace Tasks_Demo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Task.Run(OutputHello).Wait();
            return;
        }

        static void OutputHello()
        {
            Console.WriteLine("Hello!");
        }
    }
}