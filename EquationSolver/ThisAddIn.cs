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
        List<float> vars = new List<float>();
        public static int pos = 0;
        public static Node ast = new Node();
        public static String[] symbols = new String[200];
        public static String currToken = String.Empty;
        public static TreeView tr = new TreeView();
        public static int k = 0;
        private Microsoft.Office.Tools.Word.Controls.TreeView treeView = null;
        private Microsoft.Office.Tools.Word.Controls.TextBox errorTB = null;


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
                
                StringBuilder expression = new StringBuilder();
                using (XmlReader reader = XmlReader.Create(new StringReader(xmlString)))
                {
                    // XML Parsing loop
                    while (reader.Read())
                    {
                        switch (reader.Name)
                        {
                            case "m:t":
                                expression.Append(reader.ReadElementContentAsString());
                                break;
                            case "m:oMath":
                                if(reader.NodeType == XmlNodeType.EndElement)
                                    expression.Append("?");                             // end of expression
                                break;
                        }
                    }
                    if (expression.Length > 0)
                    {
                        symbols = lexic_analysis(expression.ToString());


                        Doc.Paragraphs[1].Range.InsertParagraphBefore();
                        Doc.Paragraphs[1].Range.Select();
                        Document extendedDocument = Globals.Factory.GetVstoObject(Doc);

                        if (extendedDocument.Controls.Contains(errorTB))
                        {
                            extendedDocument.Controls.Remove(errorTB);

                        }
                        errorTB = extendedDocument.Controls.AddTextBox(Doc.Paragraphs[2].Range, 100, 100, "ErrorTB");
                        if (symbols.Length > 0)
                        {
                            for (int i = 0; i < symbols.Length; i++)
                                errorTB.Text += xmlString;

                        }
                        else
                            errorTB.Text = "Ništa nije učitano" + xmlString;
                        ast = parse_string(symbols);


                        if (extendedDocument.Controls.Contains(treeView))
                        {
                            extendedDocument.Controls.Remove(treeView);

                        }
                        treeView = extendedDocument.Controls.AddTreeView(Doc.Paragraphs[1].Range, 200, 200, "EquationTR");

                        printAst(ast);
                        MessageBox.Show("Done");
                    }
                    else
                        MessageBox.Show("Nije pronađena nijedna jednadžba");


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

            int i = 0;
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
