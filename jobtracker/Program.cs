using JobTracker;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        Console.WriteLine("Job Tracker Application Starting...");
        Application.Run(new Form1());

        Console.WriteLine("Excel file created!");
    }
}