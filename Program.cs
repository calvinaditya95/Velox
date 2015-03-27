using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace LSE
{
    class Program
    {
        private static string Keyword;
        private static string rootDir;

        public struct sResult
        {
            public string path;
            public string line;
        };
        
        public static Queue<sResult> SearchResult;
        
        public class Node
        {
            public Node(string name)
            {
                this.name = name;
            }

            public string name { get; set; }
            public List<Node> Friends
            {
                get
                {
                    return FriendsList;
                }
            }

            public void isFriendOf(Node p)
            {
                FriendsList.Add(p);
            }

            List<Node> FriendsList = new List<Node>();

            public override string ToString()
            {
                return name;
            }
        }
        
        public static void PrintSearchResult()
        {
            while (SearchResult.Count > 0)
            {
                sResult temp = SearchResult.Dequeue();
                Console.WriteLine(temp.path);
                Console.WriteLine(temp.line);
                Console.WriteLine(" ");
            }
        }

        // DFS
        public static void DFS(string sDir)
        {
            try
            {
                foreach (string f in Directory.EnumerateFiles(sDir).Where(file => file.ToLower().EndsWith(".txt") || file.ToLower().EndsWith(".java") || file.ToLower().EndsWith(".html") || file.ToLower().EndsWith(".pas") || file.ToLower().EndsWith(".c") || file.ToLower().EndsWith(".cpp")))
                {
                    string[] lines = File.ReadAllLines(f);
                    foreach (string line in lines)
                    {
                        if (line.IndexOf(Keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            sResult tempSearchResult = new sResult();
                            tempSearchResult.path = f;
                            tempSearchResult.line = line;
                            SearchResult.Enqueue(tempSearchResult);
                        }
                    }
                }

                foreach (string f in Directory.EnumerateFiles(sDir).Where(file => file.ToLower().EndsWith(".doc") || file.ToLower().EndsWith(".docx")))
                {
                    Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                    object miss = System.Reflection.Missing.Value;
                    object path = f;
                    object readOnly = true;
                    Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
                    string temp;
                    string[] lines;
                    for (int i = 0; i < docs.Paragraphs.Count; i++)
                    {
                        temp = docs.Paragraphs[i + 1].Range.Text.ToString();
                        lines = temp.Split('.');
                        foreach (string line in lines)
                        {
                            if (line.IndexOf(Keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                sResult tempSearchResult = new sResult();
                                tempSearchResult.path = f;
                                tempSearchResult.line = line;
                                SearchResult.Enqueue(tempSearchResult);
                            }
                        }
                    }
                    docs.Close();
                    word.Quit();
                }

                foreach (string d in Directory.GetDirectories(sDir))
                {
                    DFS(d);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
        }

        // BFS
        public static void BFS(string sDir)
        {
            Node root = new Node(sDir);
            BuildFriendGraph(root, sDir);
            Traverse(root);
        }

        public static void BuildFriendGraph(Node s, string sDir)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    Node temp = new Node(d);
                    s.isFriendOf(temp);
                    BuildFriendGraph(temp, d);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
            }
        }

        public static void Traverse(Node root)
        {
            Queue<Node> traverseOrder = new Queue<Node>();
            Queue<Node> Q = new Queue<Node>();
            HashSet<Node> S = new HashSet<Node>();
            Q.Enqueue(root);
            S.Add(root);

            while (Q.Count > 0)
            {
                Node p = Q.Dequeue();
                traverseOrder.Enqueue(p);

                foreach (Node friend in p.Friends)
                {
                    if (!S.Contains(friend))
                    {
                        Q.Enqueue(friend);
                        S.Add(friend);
                    }
                }
            }

            while (traverseOrder.Count > 0)
            {
                Node p = traverseOrder.Dequeue();
                foreach (string f in Directory.EnumerateFiles(p.name).Where(file => file.ToLower().EndsWith(".txt") || file.ToLower().EndsWith(".java") || file.ToLower().EndsWith(".html") || file.ToLower().EndsWith(".pas") || file.ToLower().EndsWith(".c") || file.ToLower().EndsWith(".cpp")))
                {
                    string[] lines = File.ReadAllLines(f);
                    foreach (string line in lines)
                    {
                        if (line.IndexOf(Keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            sResult tempSearchResult = new sResult();
                            tempSearchResult.path = f;
                            tempSearchResult.line = line;
                            SearchResult.Enqueue(tempSearchResult);
                        }
                    }
                }

                foreach (string f in Directory.EnumerateFiles(p.name).Where(file => file.ToLower().EndsWith(".doc") || file.ToLower().EndsWith(".docx")))
                {
                    Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
                    object miss = System.Reflection.Missing.Value;
                    object path = f;
                    object readOnly = true;
                    Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
                    string temp;
                    string[] lines;
                    for (int i = 0; i < docs.Paragraphs.Count; i++)
                    {
                        temp = docs.Paragraphs[i + 1].Range.Text.ToString();
                        lines = temp.Split('.');
                        foreach (string line in lines)
                        {
                            if (line.IndexOf(Keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                sResult tempSearchResult = new sResult();
                                tempSearchResult.path = f;
                                tempSearchResult.line = line;
                                SearchResult.Enqueue(tempSearchResult);
                            }
                        }
                    }
                    docs.Close();
                    word.Quit();
                }
            }
        }

        static void Init()
        {
            Keyword = Console.ReadLine();
            rootDir = "D:\\Games";
            SearchResult = new Queue<sResult>();
        }

        static void Main(string[] args)
        {
            Init();
            DFS(rootDir);
            //BFS(rootDir);
            PrintSearchResult();
            Console.ReadLine();
        }
    }
}
