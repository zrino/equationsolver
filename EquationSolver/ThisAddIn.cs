using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;
namespace EquationSolver
{
    public partial class ThisAddIn
    {
        public static Dictionary<String,double> vars = new Dictionary<String,double>();
        public static int pos = 0;
        public static Node ast = new Node();
        public static String[] symbols = new String[200];
        public static String currToken = String.Empty;
        public static TreeView tr = new TreeView();
        public static string trName = "TreeView";
        public static TextBox tb = new TextBox();
        public static int k = 0;
        private Microsoft.Office.Tools.Word.Controls.TreeView treeView = null;
        
        public int maxEquations = 10;
        public int currentIndex = 0;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
           /* this.Application.DocumentOpen += new Word.ApplicationEvents4_DocumentOpenEventHandler(WorkWithDocument);
            ((Word.ApplicationEvents4_Event)this.Application).NewDocument +=
             new Word.ApplicationEvents4_NewDocumentEventHandler(WorkWithDocument);*/
            

            
        }
        public void WorkWithDocument()
        {
            try
            {

                Word.Document Doc = this.Application.ActiveDocument;
                string xmlString = Doc.Content.WordOpenXML;
                
                StringBuilder[] expression = new StringBuilder[maxEquations];
                for (int ix = 0; ix < maxEquations; ++ix)
                    expression[ix] = new StringBuilder();

                using (XmlReader reader = XmlReader.Create(new StringReader(xmlString)))
                {
                    // XML Parsing loop
                    while (reader.Read())
                    {
                        switch (reader.Name)
                        {
                            case "m:t":
                                expression[this.currentIndex].Append(reader.ReadElementContentAsString());
                                break;
                            case "m:oMath":
                                if (reader.NodeType == XmlNodeType.EndElement)
                                    this.currentIndex++;                 // end of expression
                                break;
                        }
                    }
                    Document extendedDocument = Globals.Factory.GetVstoObject(Doc);
                    tb = extendedDocument.Controls.AddTextBox(Doc.Paragraphs[1].Range, 200, 200, "AnswerTextBox");
                    for (int i = 0; i < currentIndex; i++)
                    {
                        if (expression[i].Length > 0)
                        {
                            pos = 0;
                            for (int k = 0; k < symbols.Length; k++)
                                symbols[k] = "";
                            lexic_analysis(expression[i].ToString()).CopyTo(symbols,0);


                            //Doc.Paragraphs[1].Range.InsertParagraphBefore();
                           // Doc.Paragraphs[1].Range.Select();
                            

                            ast = parse_string(symbols);
                            evaluateTree(ast);
                            //treeView = extendedDocument.Controls.AddTreeView(Doc.Range(25,50), 200, 200, ThisAddIn.trName + i.ToString());
                            
                            //printAst(ast);
                            
                        }
                        else
                            MessageBox.Show("Nije pronađena nijedna jednadžba");

                    }
                    foreach (KeyValuePair<string, double> entry in vars)
                    {
                        tb.Text += entry.Key.ToString() + "=" + entry.Value.ToString() + ", ";
                        // do something with entry.Value or entry.Key
                    }
                    
                    MessageBox.Show("Done");
                }

            }
            catch (Exception ex)
            {
                Word.Document currentDocument = this.Application.ActiveDocument;
                Document extendedDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
                //currentDocument.Paragraphs[1].Range.InsertParagraphBefore();
                currentDocument.Paragraphs[1].Range.InsertBefore(ex.ToString() + Environment.NewLine + symbols.ToString());
            }
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new EquationSolverRibbon();
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Word.Document currentDocument = this.Application.ActiveDocument;
            Document extendedDocument = Globals.Factory.GetVstoObject(this.Application.ActiveDocument);
            extendedDocument.Controls.Remove("EquationTR");
        }
        public static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }

        public static bool IsName(object Expression)
        {
            int counter = Regex.Matches(Expression.ToString(), @"[a-zA-Z]").Count;
            if (counter > 0)
                return true;
            return false;
        }
        public static bool IsOperator(object Expression)
        {
            String oper = Convert.ToString(Expression);
            if(oper == "="  || oper == "+" || oper == "-" || oper == "*" || oper == "/" || oper == "(" ||oper == ")")
            {
                return true;
            }
            return false;
        }
        #region Syntax abstract tree
        public static String[] lexic_analysis(String expression) //lexyc analysis, separating every character
        {

            expression = expression.Replace("+", " + ");
            expression = expression.Replace("-", " - ");
            expression = expression.Replace("=", " = ");
            expression = expression.Replace("/", " / ");
            expression = expression.Replace("*", " * ");
            expression = expression.Replace("(", " ( ");
            expression = expression.Replace(")", " ) ");
            expression = expression.Replace("±", " ± ");
            expression = expression.Replace("?", " ? ");
            char[] separators = new char[] { ' ' };
            symbols = expression.Split(separators, 100, StringSplitOptions.RemoveEmptyEntries);
            return symbols;
        }
        public void printAst(Node ast)
        {


            TreeNode ret = PrintNode(ast);
            treeView.Nodes.Add(ret);
           

        }
        public static void printTokens(String[] symbols, TextBox tb)
        {
            for (var i = 0; i < symbols.Length; i++)
            {
                tb.Text += symbols[i];
                tb.Text += " ";

            }
        }
        public static TreeNode PrintNode(Node ast)
        {
            TreeNode ret = new TreeNode();
            if (ast.op != null)
            {

                TreeNode[] trArr = new TreeNode[] { PrintNode(ast.left), PrintNode(ast.right) };
                ret = new TreeNode(ast.op, trArr);
                return ret;
            }

            else
            {
                ret = new TreeNode(ast.value.ToString());
                return ret;

            }



        }
        public static Node parse_string(String[] symbols)
        {
            Node ast = new Node();

            //int i = 0;
            ast = expression();
            return ast;
        }
        public static Node expression()
        {
            return equalityExpression();
        }
        public static Node equalityExpression()
        {
            Node left = additiveExpression();
            String tok = peekNextToken();
            while (IsOperator(tok) && (tok == "="))
            {
                skipNextToken();
                Node node = new Node();
                node.op = tok;
                node.left = left;
                node.right = additiveExpression();
                left = node;
                tok = peekNextToken();
            }
            return left;
        }
        public static Node additiveExpression()
        {
            Node left = multiplicativeExpression();
            String tok = peekNextToken();
            while (IsOperator(tok) && (tok == "+" || tok == "-"))
            {
                skipNextToken();
                Node node = new Node();
                node.op = tok;
                node.left = left;
                node.right = multiplicativeExpression();
                left = node;
                tok = peekNextToken();
            }
            return left;
        }
        public static Node multiplicativeExpression()
        {
            Node left = primaryExpression();
            String tok = peekNextToken();
            while (IsOperator(tok) && (tok == "*" || tok == "/"))
            {
                skipNextToken();
                Node node = new Node();
                node.op = tok;
                node.left = left;
                node.right = primaryExpression();
                left = node;
                tok = peekNextToken();
            }
            return left;
        }
        public static Node primaryExpression()
        {
            String tok = peekNextToken();
            if (IsNumeric(tok))
            {
                skipNextToken();
                Node node = new Node();
                node.value = tok;
                return node;
            }
            else if (IsName(tok))
            {
                skipNextToken();
                Node node = new Node();
                node.value = tok;
                return node;
            }
            else
            {
                if (IsOperator(tok) && tok == "(")
                {
                    skipNextToken();
                    Node node = expression(); //Recursion!!!
                    tok = getNextToken();
                    if (!IsOperator(tok) || tok != ")")
                    {
                        
                        throw new Exception("Error ) expected");
                    }
                    return node;
                }
                else
                {
                    
                    throw new Exception("Error " + tok + " not expected!");
                }
            }
        }
        public static String getNextToken()
        {
            String ret;
            if (currToken != String.Empty)
                ret = currToken;
            else
                ret = nextToken();
            currToken = String.Empty;
            return ret;
        }
        public static String peekNextToken()
        {
            if (currToken == String.Empty || currToken == "error")
                currToken = nextToken();

            return currToken;
        }
        public static void skipNextToken()
        {
            if (currToken == String.Empty)
                currToken = nextToken();
            currToken = String.Empty;
        }
        public static String nextToken()
        {
            String c = String.Empty;

            while (pos < symbols.Length)
            {

                c = symbols[pos++];
                if (IsNumeric(c))
                    return c;
                else if (IsOperator(c))
                    return c;
                else if (IsName(c))
                    return c;
            }
            return "error";
        }
        #endregion

        #region Evaluator
        public static double evaluateTree(Node ast)
        {
            double ret;
            if ((ast.left == null && ast.right == null) && IsNumeric(ast.value))
            {
                Double.TryParse(ast.value, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out ret);
                return ret;
            }
            else if ((ast.left == null && ast.right == null) && IsName(ast.value)) // variable is already defined
            {
                if (vars.ContainsKey(ast.value))
                    return vars[ast.value];
                else                                          //bacaj grešku
                    return 0;
            }
            else if (ast.op == "=")                            //defining var in Dictionary vars
            {
                vars[ast.left.value] = evaluateTree(ast.right);
                return 0;
            }
            else
            {
                return operate(evaluateTree(ast.left), evaluateTree(ast.right), ast.op); //else recursive call down the tree
            }
        }

        public static double operate(double arg1, double arg2, String op)
        {
            double ret;
            switch (op)
            {
                case "+":
                    ret = arg1 + arg2;
                    break;
                case "-":
                    ret = arg1 - arg2;
                    break;
                case "*":
                    ret = arg1 * arg2;
                    break;
                case "/":
                    ret = arg1 / arg2;
                    break;
                default:
                    ret = 0;
                    break;
            }
            return ret;
        }
        #endregion

    

    #region VSTO generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
    public class Node
    {
        public Node left;
        public Node right;
        public String op;
        public String value;
        public Node()
        {

        }
    }
}
