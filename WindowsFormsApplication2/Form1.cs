using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        [Serializable]
        class Node
        {
            static public SortedDictionary<myPoint, Node> dirData = new SortedDictionary<myPoint, Node>();
            
            public string data;
            public myPoint location;
            public string function;

            public List<myPoint> connetNode;
            public List<myPoint> wasConnetNode;

            /// <summary>
            /// 查询回路
            /// </summary>
            /// <param name="theNodes"></param>
            /// <param name="objNode"></param>
            /// <returns></returns>
            private bool checkCircleHelper(ref List<myPoint> theNodes,ref List<myPoint> objNode)
            {
                if (theNodes.Count == 0)
                    return true;
                var repeatList = theNodes.Intersect(objNode).ToList();
                if (repeatList.Count > 0)
                    return false;
                else
                {
                    bool result = true;
                    for (int i = 0; i < theNodes.Count; i++)
                    {
                        result &= checkCircleHelper(ref Node.dirData[theNodes[i]].wasConnetNode,ref objNode);
                    }
                    return result;
                }
            }

            /// <summary>
            /// 解析表达式,构建关系表
            /// </summary>
            /// <param name="sourceString">带待解析的字符串</param>
            /// <returns>读取数字后的表达式</returns>
            private string analysisText(string sourceString)
            {
                Regex reg = new Regex(@"R(\d+)C(\d+);");
                var mat = reg.Matches(function);

                List<myPoint> checkLoop = new List<myPoint>();
                foreach (Match item in mat)
                {
                    string t = item.Groups[0].Value;
                    int x = int.Parse(item.Groups[1].Value);
                    int y = int.Parse(item.Groups[2].Value);
                    //与数据点建立连接
                    myPoint p = new myPoint(x, y);
                    if (!dirData.ContainsKey(p))
                    {
                        dirData[p] = new Node(p);
                        dirData[p].data = "0";
                    }
                    if (!dirData[p].wasConnetNode.Contains(location))
                    {
                        dirData[p].wasConnetNode.Add(location);
                    }
                    checkLoop.Add(p);
                    sourceString = Regex.Replace(sourceString, t, dirData[p].data);
                }

                var repeatList = checkLoop.Intersect(wasConnetNode).ToList();
                connetNode = checkLoop;

                if (!checkCircleHelper(ref wasConnetNode,ref checkLoop))
                {
                    throw new Exception("存在循环引用");
                }

                return sourceString;
            }


            #region 字符串计算器
            string nums = "N0123456789.";
            string symbols = "-+*/%()";

            /// <summary>
            /// 处理表达式中多余的正号和负号
            /// </summary>
            /// <param name="sourceString">待处理的表达式</param>
            /// <returns>处理后的表达式</returns>
            private string dealWithString(string sourceString)
            {
                string result = "";
                int n = sourceString.Count();
                int m = 0;
                if(n == 0)
                {
                    throw new Exception("字符串为空");
                }
                while (n != m)
                {
                    n = sourceString.Count();
                    result = "";
                    for (int i = 0; i < sourceString.Length; ++i)
                    {
                        if (sourceString[i] == '-')
                        {
                            if (sourceString[i + 1] == '-')
                            {
                                i += 1;
                                result += "+";
                                continue;
                            }
                            if (i == 0 || sourceString[i - 1] == '+' || sourceString[i - 1] == '-'
                                || sourceString[i - 1] == '*' || sourceString[i - 1] == '/'
                                || sourceString[i - 1] == '%' || sourceString[i - 1] == '(')
                            {
                                result += 'N';
                            }
                            else
                                result += '-';
                        }
                        else if (sourceString[i] == '+')
                        {
                            if (i == 0 || sourceString[i - 1] == '+' || sourceString[i - 1] == '-'
                                || sourceString[i - 1] == '*' || sourceString[i - 1] == '/'
                                || sourceString[i - 1] == '%' || sourceString[i - 1] == '(')
                            {
                            }
                            else
                                result += '+';
                        }
                        else
                        {
                            result += sourceString[i];
                        }
                    }
                    m = result.Count();
                    sourceString = result;
                }
                return result;
            }

            /// <summary>
            /// 获取操作符优先级
            /// </summary>
            /// <param name="op">操作符</param>
            /// <returns>优先级</returns>
            private int prior(char op)
            {
                if (op == '+' || op == '-')
                    return 1;
                if (op == '*' || op == '/' || op == '%')
                    return 2;
                return 0;
            }

            /// <summary>
            /// 进行双目运算并将结果存入栈中
            /// </summary>
            /// <param name="symbol">操作符</param>
            /// <param name="n1">左操作数</param>
            /// <param name="n2">右操作数</param>
            /// <param name="numsStack">数字栈</param>
            private void cal(char symbol, double n1, double n2, ref Stack<double> numsStack)
            {
                switch (symbol)
                {
                    case '+':
                        numsStack.Push(n1 + n2);
                        break;
                    case '-':
                        numsStack.Push(n1 - n2);
                        break;
                    case '*':
                        numsStack.Push(n1 * n2);
                        break;
                    case '/':
                        if(n2 == 0)
                        {
                            throw new Exception("除数不能为0");
                        }
                        numsStack.Push(n1 / n2);

                        break;
                    case '%':
                        numsStack.Push(n1 % n2);
                        break;
                }
            }

            /// <summary>
            /// 从字符串中提取数字到栈中
            /// </summary>
            /// <param name="sourceString">源字符串</param>
            /// <param name="numsStack">数字栈</param>
            private void getNumber(ref string sourceString, ref Stack<double> numsStack)
            {
                int b = 0;
                string num = "";
                double tryParse = 0;
                for (; b < sourceString.Length; ++b)
                {
                    if (!nums.Contains(sourceString[b]))
                    {
                        break;
                    }
                    else
                    {
                        if (sourceString[b] == 'N')
                        {
                            num += '-';
                        }
                        else
                        {
                            num += sourceString[b];
                        }
                    }
                }
                if (double.TryParse(num, out tryParse))
                {
                    numsStack.Push(tryParse);
                }
                sourceString = sourceString.Substring(b);
            }

            /// <summary>
            /// 从字符串中提取运算符到栈中
            /// </summary>
            /// <param name="sourceString">源字符串</param>
            /// <param name="symbolStack">数字栈</param>
            private void getSymbol(ref string sourceString,ref Stack<char> symbolStack)
            {
                symbolStack.Push(sourceString[0]);
                sourceString = sourceString.Substring(1);
            }

            /// <summary>
            /// 去除运算公式中的括号
            /// </summary>
            /// <param name="sourceString">源字符串</param>
            /// <returns>去除括号后并计算的运算公式</returns>
            private string removeParentheses(string sourceString)
            {
                if(sourceString.IndexOf('(')==-1)
                {
                    return calculater(dealWithString(sourceString));
                }
                Regex reg = new Regex(@"\(([0-9\.\+\-\*/%]*?)\)");
                string turnSymbol = "+-*.";
                string pattern = "";
                while (sourceString.IndexOf('(') != -1)
                {
                    var item = reg.Match(sourceString);
                    string result = "";
                    result = removeParentheses(item.Groups[1].Value);
                    for (int i = 0; i < item.Groups[1].Value.Length; ++i)
                    {
                        if (turnSymbol.Contains(item.Groups[1].Value[i]))
                        {
                            pattern += "\\" + item.Groups[1].Value[i];
                        }
                        else
                        {
                            pattern += item.Groups[1].Value[i];
                        }
                    }
                    sourceString = Regex.Replace(sourceString, "\\(" + pattern + "\\)", result);
                    pattern = "";
                }
                return sourceString;
            }

            /// <summary>
            /// 进行多数多运算符表达式的计算
            /// </summary>
            /// <param name="exp">表达式</param>
            /// <returns>计算结果</returns>
            private string calculater(string exp)
            {
                Stack<double> numsStack = new Stack<double>();
                Stack<char> symbolStack = new Stack<char>();


                bool couldcalculate = true;
                while (exp.Length > 0)
                {
                    char newSymbol, oldSymbol;
                    if (symbolStack.Count > 1 && couldcalculate)
                    {
                        newSymbol = symbolStack.Pop();
                        oldSymbol = symbolStack.Peek();

                        if (prior(newSymbol) <= prior(oldSymbol))
                        {
                            double n2 = numsStack.Pop(), n1 = numsStack.Pop();
                            symbolStack.Pop();
                            cal(oldSymbol, n1, n2, ref numsStack);
                            symbolStack.Push(newSymbol);
                        }
                        else
                        {
                            symbolStack.Push(newSymbol);
                            couldcalculate = false;
                        }
                    }
                    //获取运算符
                    else if (symbols.Contains(exp[0]))
                    {
                        getSymbol(ref exp, ref symbolStack);
                        if (symbolStack.Peek() != '(')
                        {
                            couldcalculate = true;
                        }
                    }
                    //获取数字
                    else if (nums.Contains(exp[0]))
                    {
                        getNumber(ref exp, ref numsStack);
                    }
                }

                while (symbolStack.Count != 0)
                {
                    double n2 = numsStack.Pop(), n1 = numsStack.Pop();
                    char oldSymbol = symbolStack.Pop();
                    cal(oldSymbol, n1, n2, ref numsStack);
                }
                return numsStack.Pop().ToString();
            }

            /// <summary>
            /// 字符串运算的总接口
            /// </summary>
            /// <param name="text">待运算的表达式</param>
            private void calculation(string text)
            {
                //去除公式中的空格
                string sourceString = text.Trim();
                function = "";
                for (int i = 0; i < sourceString.Length; ++i)
                {
                    if (sourceString[i] != ' ')
                    {
                        function += sourceString[i];
                    }
                }
                sourceString = analysisText(function);
                //如果算式中包含非法字符,则显示ERROR!
                for(int i = 0;i<sourceString.Length;++i)
                {
                    if(!nums.Contains(sourceString[i])&&!symbols.Contains(sourceString[i]))
                    {
                        data = "ERROR!";
                        return;
                    }
                }

                data = calculater(dealWithString(removeParentheses(sourceString)));

            }
            #endregion


            /// <summary>
            /// 解除绑定
            /// </summary>
            private void disconnect()
            {
                if (connetNode.Count != 0)
                {
                    for (int i = 0; i < connetNode.Count; i++)
                    {
                        dirData[connetNode[i]].wasConnetNode.Remove(location);
                    }
                }
            }

            /// <summary>
            /// 广播自己已经被改变的消息,通知与之关联的节点重新计算数值
            /// </summary>
            private void notice()
            {
                for(int i = 0;i<wasConnetNode.Count;++i)
                { 
                    Node temp = dirData[wasConnetNode[i]];
                    temp.calculation(temp.function);
                    temp.notice();
                }
            }

            /// <summary>
            /// 修改信息的总接口
            /// </summary>
            /// <param name="text">源表达式</param>
            public void write(string text,TextBox txb = null)
            {
                try
                {
                    if (text.IndexOf('=') == 0)
                    {
                        try
                        {
                            disconnect();
                            calculation(text.Substring(1));
                            notice();
                        }
                        catch(Exception e)
                        {
                            MessageBox.Show("公式错误:" + e.Message);
                            string preData = this.data;
                            this.write(preData);
                            if (txb != null)
                            {
                                txb.Text = preData;
                            }
                        }

                    }
                    else
                    {
                        data = text;
                        disconnect();
                        connetNode.Clear();
                        function = "";
                        notice();
                    }
                }
                catch(Exception e)
                {
                    MessageBox.Show("请检查输入是否正确\n错误信息:" + e.Message);
                    string preData = this.data;
                    this.write(preData);
                    if (txb != null)
                    {
                        txb.Text = preData;
                    }
                }
            }

            public Node(myPoint point)
            {
                location = point;
                data = "";
                function = "";
                connetNode = new List<myPoint>();
                wasConnetNode = new List<myPoint>();
            }

        }
        [Serializable]
        class myPoint:IComparable<myPoint>
        {
            public int X;
            public int Y;
            public myPoint(int x,int y)
            {
                X = x;
                Y = y;
            }
            
            public int CompareTo(myPoint other)
            {
                if (this.X == other.X)
                {
                    if (this.Y == other.Y)
                    {
                        return 0;
                    }
                    else if (this.Y > other.Y)
                    {
                        return 1;
                    }
                    else
                    {
                        return -1;
                    }
                }
                else if (this.X > other.X)
                {
                    return 1;
                }
                else
                {
                    return -1;
                }
            }

            public override bool Equals(object obj)
            {
                if (obj == null)
                {
                    return false;
                }
                if ((obj.GetType().Equals(this.GetType())) == false)
                {
                    return false;
                }
                myPoint temp =(myPoint)obj;

                return this.X.Equals(temp.X) && this.Y.Equals(temp.Y);

            }

            //重写GetHashCode方法（重写Equals方法必须重写GetHashCode方法，否则发生警告

            public override int GetHashCode()
            {
                return this.X.GetHashCode() + this.Y.GetHashCode();
            }
        }

        #region 界面

        List<Button> colButton = new List<Button>();
        List<Button> rowButton = new List<Button>();
        Panel pan;
        PictureBox picbox;
        Graphics g;
        Image img;
        //记录最左上角的坐标
        Point p = new Point(0, 0);
        //记录被选中的坐标
        myPoint selectPoint = null;
        myPoint expModPoint = null;
        VScrollBar vsb;
        HScrollBar hsb;
        //记录是否进入表达式模式
        private void create()
        {
            pan = new Panel();
            this.Controls.Add(pan);
            vsb = new VScrollBar();
            vsb.Dock = DockStyle.Right;
            hsb = new HScrollBar();
            hsb.Dock = DockStyle.Bottom;
            vsb.Maximum = 65536;
            vsb.Minimum = 0;
            hsb.Maximum = 256;
            hsb.Minimum = 0;
            

            vsb.ValueChanged += Vsb_ValueChanged;
            hsb.ValueChanged += Hsb_ValueChanged;

            
            pan.Controls.Add(vsb);
            pan.Controls.Add(hsb);
            pan.Location = new Point(0, 200);
            pan.Size = new Size(this.Size.Width-20, this.Size.Height - 240);
            Button fstBtn = new Button();
            fstBtn.Location = new Point(0, 0);
            fstBtn.Size = new Size(50, 20);
            pan.Controls.Add(fstBtn);
            int colWidth = 75;
            int colNum = pan.Width / colWidth;
            int rowHight = 20;
            int rowNum = pan.Height / rowHight;
            //生成列按钮
            for (int i = 0; i < colNum; ++i)
            {
                Button btn = new Button();
                btn.Tag = i;
                btn.Text = "C" + (i + 1);
                btn.Location = new Point(50 + i * 75, 0);
                btn.Size = new Size(75, 20);
                pan.Controls.Add(btn);
                colButton.Add(btn);
            }
            //生成行按钮
            for (int i = 0; i < rowNum; ++i)
            {
                Button btn = new Button();
                btn.Tag = i;
                btn.Text = "R" + (i + 1);
                btn.Location = new Point(0, 20 + i * 20);
                btn.Size = new Size(50, 20);
                btn.Enter += Btn_Enter;
                pan.Controls.Add(btn);
                rowButton.Add(btn);
            }
            //生成画板,用于绘制表格
            picbox = new PictureBox();
            picbox.Location = new Point(50, 20);
            picbox.Size = new Size(pan.Width- 50, pan.Height - 20);
            pan.Controls.Add(picbox);
            picbox.MouseClick += Picbox_MouseClick;
            picbox.MouseDoubleClick += Picbox_MouseDoubleClick;
            picbox.PreviewKeyDown += Picbox_PreviewKeyDown;
            picbox.MouseWheel += Picbox_MouseWheel;
            hsb.Enter += Hsb_Enter;
            vsb.Enter += Vsb_Enter;
        }

        private void Vsb_Enter(object sender, EventArgs e)
        {
            picbox.Focus();
        }

        private void Hsb_Enter(object sender, EventArgs e)
        {
            picbox.Focus();
        }

        private void Btn_Enter(object sender, EventArgs e)
        {
            picbox.Focus();
        }

        private void Picbox_MouseWheel(object sender, MouseEventArgs e)
        {
            if (e.Delta < 0&& vsb.Value<65535)
            {
                vsb.Value++;
                picbox.Focus();
                p.X = vsb.Value;
                for (int i = 0; i < rowButton.Count; ++i)
                {
                    rowButton[i].Text = "R" + (i + vsb.Value + 1);
                }

            }
            else if(e.Delta > 0 && vsb.Value >0)
            {
                vsb.Value--;
                picbox.Focus();
                p.X = vsb.Value;
                for (int i = 0; i < rowButton.Count; ++i)
                {
                    rowButton[i].Text = "R" + (i + vsb.Value + 1);
                }
            }

            drawPic();
            drawRectangle(selectPoint,Color.Black);
            drawSelectRect(selectPoints);
        }
        
        private void Hsb_ValueChanged(object sender, EventArgs e)
        {
            picbox.Focus();
            var hsbar = sender as HScrollBar;
            p.Y = hsbar.Value;
            drawPic();
            for (int i = 0; i < colButton.Count; ++i)
            {
                colButton[i].Text = "C" + (i + hsbar.Value + 1);
            }
            drawRectangle(selectPoint,Color.Black);
            drawSelectRect(selectPoints);
        }

        private void Vsb_ValueChanged(object sender, EventArgs e)
        {
            picbox.Focus();
            var vsbar = sender as VScrollBar;
            p.X = vsbar.Value;
            drawPic();
            for (int i = 0; i < rowButton.Count; ++i)
            {
                rowButton[i].Text = "R" + (i + vsbar.Value+1);
            }
            drawRectangle(selectPoint,Color.Black);
            drawSelectRect(selectPoints);
        }

        //双击某绘制的单元格处会在此处生成一个textbox,用于输入文本
        private void Picbox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if(!isExpModing)
                generateTextBox(selectPoint);
        }
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Shift & e.KeyCode == Keys.Oemplus)
            {
                
            }
            else if (e.KeyCode == Keys.Oemplus && selectPoint != null)
            {
                selectPoints = "";
                isExpModing = true;
                expModPoint = selectPoint;
            }
        }

        myPoint expSelectPoint = null;

        string selectPoints = "";
        private void Picbox_MouseClick(object sender, MouseEventArgs e)
        {
            picbox.Focus();
            int x = e.X / 75;
            int y = e.Y / 20;
            if (!isExpModing)
            {
                selectPoint = new myPoint(p.X + y, p.Y + x);
                drawPic();
                g.DrawRectangle(new Pen(Color.Black, 2), (selectPoint.Y - p.Y) * 75, (selectPoint.X - p.X) * 20, 75, 20);
                picbox.Image = img;
                if (Node.dirData.ContainsKey(selectPoint))
                {
                    if (Node.dirData[selectPoint].function != "")
                    {
                        textBox1.Text = "=" + Node.dirData[selectPoint].function;
                        drawSelectRect(textBox1.Text);

                    }
                    else
                        textBox1.Text = Node.dirData[selectPoint].data;
                    textBox1.Tag = selectPoint;
                }else
                {
                    textBox1.Text = "";
                }
            }
            else
            {
                if (isExpModing)
                {
                    expSelectPoint = new myPoint(p.X + y, p.Y + x);
                    string preString = textBox1.Text;
                    string symbols = "-+*/%(";
                    if (symbols.Contains(preString[preString.Length-1]))
                        textBox1.Text += "R" + expSelectPoint.X + "C" + expSelectPoint.Y + ";";
                    else
                    {
                        int index = preString.LastIndexOf('R');
                        if (index != -1)
                        {
                            preString = preString.Substring(0, index);
                            index = selectPoints.LastIndexOf('R');
                            selectPoints = selectPoints.Substring(0, index);
                        }
                        textBox1.Text = preString+ "R" + expSelectPoint.X + "C" + expSelectPoint.Y + ";";
                    }
                    selectPoints += "R" + expSelectPoint.X + "C" + expSelectPoint.Y + ";,";
                    drawPic();
                    g.DrawRectangle(new Pen(Color.Black, 2), (selectPoint.Y - p.Y) * 75, (selectPoint.X - p.X) * 20, 75, 20);
                    drawSelectRect(selectPoints);
                }
            }
        }

        private void drawSelectRect(string data)
        {
            Regex reg = new Regex(@"R(\d+)C(\d+);");
            var mat = reg.Matches(textBox1.Text);
            
            foreach (Match item in mat)
            {
                int m = int.Parse(item.Groups[1].Value);
                int n = int.Parse(item.Groups[2].Value);
                drawRectangle(new myPoint(m, n), Color.Orange);
            }

        }

        private void generateTextBox(myPoint loc)
        {
            if (loc.X >= p.X && loc.Y >= p.Y&& p.X + rowButton.Count - 1 > selectPoint.X + 1)
            {
                #region 生成文本框
                TextBox txb = new TextBox();
                //使其坐标变为整数(取左上角)
                int x = (loc.X - p.X) * 20;
                int y = (loc.Y - p.Y) * 75;
                txb.Location = new Point(picbox.Location.X + y, picbox.Location.Y + x);
                txb.Size = new Size(75, 20);
                pan.Controls.Add(txb);
                txb.BringToFront();
                txb.Focus();
                #endregion

                //记录此时文本框的信息
                int m = x / 20;
                int n = y / 75;
                var tp = new myPoint(p.X + m, p.Y + n);
                if (Node.dirData.ContainsKey(tp))
                {
                    if (Node.dirData[tp].function != "")
                        txb.Text = "=" + Node.dirData[tp].function;
                    else
                        txb.Text = Node.dirData[tp].data;
                }
                txb.Tag = new Point((int)rowButton[m].Tag, (int)colButton[n].Tag);
                txb.Leave += Txb_Leave;
                txb.PreviewKeyDown += Txb_PreviewKeyDown;
                txb.TextChanged += Txb_TextChanged;
            }
        }

        private void Txb_TextChanged(object sender, EventArgs e)
        {

            textBox1.Text = ((TextBox)sender).Text;
        }

        private void Txb_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Oemplus && selectPoint != null)
            {
                selectPoints = "";
                isExpModing = true;
                expModPoint = selectPoint;
            }
            if (e.KeyCode == Keys.Enter)
            {
                picbox.Focus();
            }
            drawRectangle(selectPoint,Color.Black);
        }

        Thread thread;

        private void Form1_Load(object sender, EventArgs e)
        {

            thread = new Thread(new ThreadStart(autoSave));
            thread.Start();
            create();
        }

        private void drawRectangle(myPoint point,Color col)
        {
            if (point != null)
            {
                if (point.X >= p.X && point.Y >= p.Y)
                {
                    g.DrawRectangle(new Pen(col, 2), (point.Y - p.Y) * 75, (point.X - p.X) * 20, 75, 20);
                    picbox.Image = img;
                }
            }
        }

        private void Picbox_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            switch(e.KeyCode)
            {
                case Keys.Oemplus:
                    {
                        if (!isExpModing)
                        {
                            if (selectPoint != null)
                            {
                                selectPoints = "";
                                isExpModing = true;
                                expModPoint = selectPoint;
                                textBox1.Text = "=";
                            }
                        }
                        break;
                    }
                case Keys.Enter:
                    {
                        if(selectPoint!=null)
                        {
                            generateTextBox(selectPoint);
                        }
                        if(isExpModing)
                        {
                            isExpModing = false;
                        }
                        break;
                    }
                case Keys.Up:
                    {
                        if (p.X < selectPoint.X)
                        {
                            selectPoint.X--;
                        }
                        else
                        {
                            if (vsb.Value > 0)
                            {
                                vsb.Value--;
                                selectPoint.X--;
                                picbox.Focus();
                                p.X = vsb.Value;
                                drawPic();
                                for (int i = 0; i < rowButton.Count; ++i)
                                {
                                    rowButton[i].Text = "R" + (i + vsb.Value + 1);
                                }
                            }
                        }
                        if (Node.dirData.ContainsKey(selectPoint))
                        {
                            if (Node.dirData[selectPoint].function != "")
                                textBox1.Text = "=" + Node.dirData[selectPoint].function;
                            else
                                textBox1.Text = Node.dirData[selectPoint].data;
                        }
                        else
                        {
                            textBox1.Text = "";
                        }
                        drawPic();
                        drawRectangle(selectPoint,Color.Black);
                        drawSelectRect(textBox1.Text);
                        picbox.Focus();
                        break;
                    }
                case Keys.Down:
                    {
                        if (p.X+rowButton.Count-2 > selectPoint.X+1)
                        {
                            selectPoint.X++;
                        }
                        else
                        {
                            if (vsb.Value < 65535)
                            {
                                vsb.Value++;
                                selectPoint.X++;
                                picbox.Focus();
                                p.X = vsb.Value;
                                for (int i = 0; i < rowButton.Count; ++i)
                                {
                                    rowButton[i].Text = "R" + (i + vsb.Value + 1);
                                }
                            }
                        }
                        if (Node.dirData.ContainsKey(selectPoint))
                        {
                            if (Node.dirData[selectPoint].function != "")
                                textBox1.Text = "=" + Node.dirData[selectPoint].function;
                            else
                                textBox1.Text = Node.dirData[selectPoint].data;
                        }
                        else
                        {
                            textBox1.Text = "";
                        }
                        drawPic();
                        drawRectangle(selectPoint,Color.Black);
                        drawSelectRect(textBox1.Text);
                        picbox.Focus();
                        break;
                    }
                case Keys.Left:
                    {
                        if (p.Y < selectPoint.Y)
                        {
                            selectPoint.Y--;
                            picbox.Focus();
                        }
                        else
                        {
                            if (hsb.Value > 0)
                            {
                                hsb.Value--;
                                selectPoint.Y--;
                                picbox.Focus();
                                p.Y = hsb.Value;
                                for (int i = 0; i < colButton.Count; ++i)
                                {
                                    colButton[i].Text = "R" + (i + hsb.Value + 1);
                                }
                            }
                        }
                        if (Node.dirData.ContainsKey(selectPoint))
                        {
                            if (Node.dirData[selectPoint].function != "")
                                textBox1.Text = "=" + Node.dirData[selectPoint].function;
                            else
                                textBox1.Text = Node.dirData[selectPoint].data;
                        }
                        else
                        {
                            textBox1.Text = "";
                        }
                        drawPic();
                        drawRectangle(selectPoint,Color.Black);
                        drawSelectRect(textBox1.Text);
                        picbox.Focus();
                        break;
                    }

                case Keys.Right:
                    {
                        if (p.Y + colButton.Count > selectPoint.Y + 2)
                        {
                            selectPoint.Y++;
                        }
                        else
                        {
                            if (hsb.Value < 256)
                            {
                                hsb.Value++;
                                selectPoint.Y++;
                                picbox.Focus();
                                p.Y = hsb.Value;
                                for (int i = 0; i < colButton.Count; ++i)
                                {
                                    colButton[i].Text = "R" + (i + hsb.Value + 1);
                                }
                            }
                        }
                        if (Node.dirData.ContainsKey(selectPoint))
                        {
                            if (Node.dirData[selectPoint].function != "")
                                textBox1.Text = "="+Node.dirData[selectPoint].function;
                            else
                                textBox1.Text = Node.dirData[selectPoint].data;
                        }
                        else
                        {
                            textBox1.Text = "";
                        }
                        drawPic();
                        drawRectangle(selectPoint,Color.Black);
                        drawSelectRect(textBox1.Text);
                        picbox.Focus();
                        break;
                    }
            }
        }

        private void autoSave()
        {
            while(true)
            {
                Thread.Sleep(300000);
                save("autoSave.txt");
            }
        }

        private bool isEmptyNode(Node node)
        {
            if (node.data == "" && node.function == "" && node.connetNode.Count == 0 && node.wasConnetNode.Count == 0)
                return true;
            else
                return false;

        }

        private void Txb_Leave(object sender, EventArgs e)
        {
            var txb = sender as TextBox;
            if (!isExpModing)
            {
                
                myPoint point = new myPoint(p.X + ((Point)txb.Tag).X, p.Y + ((Point)txb.Tag).Y);
                if (!Node.dirData.ContainsKey(point))
                {
                    Node.dirData[point] = new Node(point);
                }
                Node.dirData[point].write(txb.Text.Trim());
                drawPic();
            }
            txb.Dispose();
        }

        private void drawPic()
        {
            img = new Bitmap(picbox.Width, picbox.Height);
            g = Graphics.FromImage(img);
            g.Clear(Color.White);
            
            Pen pen = new Pen(System.Drawing.ColorTranslator.FromHtml("#D4D4D4"), 1);
            Point p1 = new Point(75, 0);
            Point p2 = new Point(75, picbox.Height);
            for (int i = 0; i < colButton.Count; ++i)
            {
                g.DrawLine(pen, p1, p2);
                p1.X += 75;
                p2.X = p1.X;
            }
            p1.X = 0;p1.Y = 20;
            p2.X = picbox.Width;p2.Y = 20;
            for (int i = 0;i<rowButton.Count;++i)
            {
                g.DrawLine(pen, p1, p2);
                p1.Y += 20;
                p2.Y = p1.Y;
            }
            for(int i = 0;i<rowButton.Count;++i)
            {
                for(int j = 0;j<colButton.Count;++j)
                {
                    myPoint tempPoint = new myPoint(p.X + i, p.Y + j);
                    if (Node.dirData.ContainsKey(tempPoint))
                    {
                        Font font = new Font("Consolas", 10);
                        Brush brush = Brushes.Black;
                        {
                            if (Node.dirData[tempPoint].data.Length > 5)
                                g.DrawString(Node.dirData[tempPoint].data.Substring(0, 5) + "...", font, brush, j * 75, i * 20);
                            else
                                g.DrawString(Node.dirData[tempPoint].data, font, brush, j * 75, i * 20);
                        }
                    }
                }
            }
            picbox.Image = img;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            colButton.Clear();
            rowButton.Clear();
            pan.Dispose();
            create();
        }

        private void Form1_Paint(object sender, PaintEventArgs e)
        {
            drawPic();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            thread.Abort();
        }
        bool isExpModing = false;
        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (!isExpModing)
            {
                var sp = (myPoint)textBox1.Tag;
                if (sp != null)
                {
                    Node.dirData[sp].write(textBox1.Text, textBox1);
                }
            }
        }

        private void textBox1_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            

            if (e.KeyCode == Keys.Enter)
            {
                if (isExpModing)
                {
                    isExpModing = false;
                    if (!Node.dirData.ContainsKey(expModPoint))
                    {
                        Node.dirData[expModPoint] = new Node(expModPoint);
                    }
                    Node.dirData[expModPoint].write(textBox1.Text, textBox1);

                    drawPic();
                    expSelectPoint = null;
                    expSelectPoint = null;
                }
                else
                {
                    if (selectPoint != null)
                    {
                        if (!Node.dirData.ContainsKey(selectPoint))
                        {
                            Node.dirData[selectPoint] = new Node(selectPoint);
                        }
                        Node.dirData[selectPoint].write(textBox1.Text, textBox1);
                        drawPic();
                    }
                    else
                    {
                        MessageBox.Show("未选中单元格");
                    }
                }
            }
        }
        #endregion

        #region 读取保存

        /// <summary>
        /// 解压
        /// </summary>
        /// <param name="Value"></param>
        /// <returns></returns>
        public static DataSet GetDatasetByString(string Value)
        {
            DataSet ds = new DataSet();
            string CC = GZipDecompressString(Value);
            System.IO.StringReader Sr = new StringReader(CC);
            ds.ReadXml(Sr);
            return ds;
        }

        /// <summary>
        /// 根据DATASET压缩字符串
        /// </summary>
        /// <param name="ds"></param>
        /// <returns></returns>
        public static string GetStringByDataset(string ds)
        {
            return GZipCompressString(ds);
        }

        /// <summary>
        /// 将传入字符串以GZip算法压缩后，返回Base64编码字符
        /// </summary>
        /// <param name="rawString">需要压缩的字符串</param>
        /// <returns>压缩后的Base64编码的字符串</returns>
        public static string GZipCompressString(string rawString)
        {
            if (string.IsNullOrEmpty(rawString) || rawString.Length == 0)
            {
                return "";
            }
            else
            {
                byte[] rawData = System.Text.Encoding.UTF8.GetBytes(rawString.ToString());
                byte[] zippedData = Compress(rawData);
                return (string)(Convert.ToBase64String(zippedData));
            }
        }

        /// <summary>
        /// GZip压缩
        /// </summary>
        /// <param name="rawData"></param>
        /// <returns></returns>
        static byte[] Compress(byte[] rawData)
        {
            MemoryStream ms = new MemoryStream();
            GZipStream compressedzipStream = new GZipStream(ms, CompressionMode.Compress, true);
            compressedzipStream.Write(rawData, 0, rawData.Length);
            compressedzipStream.Close();
            return ms.ToArray();
        }


        /// <summary>
        /// 将传入的二进制字符串资料以GZip算法解压缩
        /// </summary>
        /// <param name="zippedString">经GZip压缩后的二进制字符串</param>
        /// <returns>原始未压缩字符串</returns>
        public static string GZipDecompressString(string zippedString)
        {
            if (string.IsNullOrEmpty(zippedString) || zippedString.Length == 0)
            {
                return "";
            }
            else
            {
                byte[] zippedData = Convert.FromBase64String(zippedString.ToString());
                return (string)(System.Text.Encoding.UTF8.GetString(Decompress(zippedData)));
            }
        }


        /// <summary>
        /// ZIP解压
        /// </summary>
        /// <param name="zippedData"></param>
        /// <returns></returns>
        public static byte[] Decompress(byte[] zippedData)
        {
            MemoryStream ms = new MemoryStream(zippedData);
            GZipStream compressedzipStream = new GZipStream(ms, CompressionMode.Decompress);
            MemoryStream outBuffer = new MemoryStream();
            byte[] block = new byte[1024];
            while (true)
            {
                int bytesRead = compressedzipStream.Read(block, 0, block.Length);
                if (bytesRead <= 0)
                    break;
                else
                    outBuffer.Write(block, 0, bytesRead);
            }
            compressedzipStream.Close();
            return outBuffer.ToArray();
        }


        private void save(string fileName)
        {
            FileStream fs = new FileStream(fileName,FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            StringBuilder data = new StringBuilder("");
            foreach (var item in Node.dirData)
            {
                myPoint p = item.Key;
                Node n = item.Value;
                string helper = n.function + "|" + "R" + p.X.ToString() + "C" + p.Y.ToString() + ";|";
                foreach (var con in n.connetNode)
                {
                    helper += "R" + con.X.ToString() + "C" + con.Y.ToString() + ";";
                }
                helper += "|";

                foreach (var con in n.wasConnetNode)
                {
                    helper += "R" + con.X.ToString() + "C" + con.Y.ToString() + ";";
                }
                helper += "|"+n.data;
                data.Append(helper + "\r\n");
            }
            string result = GZipCompressString(data.ToString());
            sw.Write(result);
            sw.Close();
            fs.Close();

            MessageBox.Show("保存成功");
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            read();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "txt文件(*.txt) | *.txt";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                save(sfd.FileName);
            }
        }
        private string[] split(string obj)
        {
            string[] result = new string[5];
            for (int i = 0; i < 4; i++)
            {
                int index = obj.IndexOf('|');
                result[i] = obj.Substring(0,index);
                obj = obj.Substring(index+1);
            }
            result[4] = obj;
            return result;
        }
        private void read()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "txt文件(*.txt) | *.txt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate);
                StreamReader sr = new StreamReader(fs);
                string data = sr.ReadToEnd();

                string result = GZipDecompressString(data);

                MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(result));
                stream.Position = 0;
                sr = new StreamReader(stream);
                Node.dirData.Clear();
                string helper = sr.ReadLine();
                Regex reg = new Regex(@"R(\d+)C(\d+);");

                while (helper != null)
                {
                    var info = split(helper);
                    var match = reg.Match(info[1]);
                    int x = int.Parse(match.Groups[1].Value);
                    int y = int.Parse(match.Groups[2].Value);
                    myPoint p = new myPoint(x, y);
                    Node n = new Node(p);
                    n.data = info[4];
                    n.function = info[0];
                    n.location = p;
                    var mat = reg.Matches(info[2]);
                    foreach (Match item in mat)
                    {
                        x = int.Parse(item.Groups[1].Value);
                        y = int.Parse(item.Groups[2].Value);
                        n.connetNode.Add(new myPoint(x, y));
                    }
                    mat = reg.Matches(info[3]);
                    foreach (Match item in mat)
                    {
                        x = int.Parse(item.Groups[1].Value);
                        y = int.Parse(item.Groups[2].Value);
                        n.wasConnetNode.Add(new myPoint(x, y));
                    }
                    Node.dirData[p] = n;
                    helper = sr.ReadLine();
                }
                drawPic();
                
                sr.Close();
                fs.Close();
                
                stream.Close();
                isExpModing = false;
                expModPoint = null;
                selectPoint = null;
                selectPoints = "";
                textBox1.Text = "";

                MessageBox.Show("读取成功");
                
            }

        }
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "csv文件(*.csv) | *.csv";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = new FileStream(ofd.FileName, FileMode.OpenOrCreate);
                StreamReader sr = new StreamReader(fs);

                Node.dirData.Clear();
                string helper = sr.ReadLine();

                int r = 0, c = 0;
                while (helper != null)
                {

                    var nums = helper.Split(',');
                    c = 0;
                    while (c < nums.Count())
                    {
                        if (nums[c] != "")
                        {
                            myPoint mp = new myPoint(r, c);
                            Node.dirData[mp] = new Node(mp);
                            Node.dirData[mp].write(nums[c]);
                        }
                        c++;
                    }
                    helper = sr.ReadLine();
                    r++;
                }
                drawPic();
                sr.Close();
                fs.Close();
            }

            isExpModing = false;
            expModPoint = null;
            selectPoint = null;
            selectPoints = "";
            textBox1.Text = "";
        }


        #endregion


    }
}