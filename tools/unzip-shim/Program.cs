using System;
using System.IO;
using System.IO.Compression;

public static class Program
{
    public static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.Error.WriteLine("Usage: unzip -Z1 archive | unzip -p archive entry");
            return 2;
        }

        using (ZipArchive archive = ZipFile.OpenRead(args[1]))
        {
            if (args[0] == "-Z1")
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                    Console.WriteLine(entry.FullName);
                return 0;
            }

            if (args[0] == "-p" && args.Length >= 3)
            {
                ZipArchiveEntry entry = archive.GetEntry(args[2]);
                if (entry == null) return 11;
                using (Stream input = entry.Open())
                using (Stream output = Console.OpenStandardOutput())
                    input.CopyTo(output);
                return 0;
            }
        }

        Console.Error.WriteLine("Unsupported unzip arguments");
        return 2;
    }
}
