using WinExplorer;

namespace ExplorerShellTestWinFormsApp
{
    public partial class Form1 : Form
    {
        WinExplorerLikeContextMenu? _explorerLikeContextMenu;

        public Form1()
        {
            InitializeComponent();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            _explorerLikeContextMenu = new WinExplorerLikeContextMenu();
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            base.OnHandleDestroyed(e);

            _explorerLikeContextMenu?.Dispose();
            _explorerLikeContextMenu = null;
        }

        private void Form1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
            {
                return;
            }

            if (_explorerLikeContextMenu is null)
            {
                return;
            }

            var fileInfos = new FileSystemInfo[]
            {
                new FileInfo(@"C:\var\log\AutomaticDisposeGenerator\20220921110613_18312.txt"),
                //new FileInfo(@"C:\var\log\AutomaticNotifyPropertyChangedImpl\20220921110613_18312.txt"),
                //new FileInfo(@"C:\var\log\AutomaticNotifyPropertyChangedImpl\20220921110623_6368.txt"),
            };

            var screenPoint = PointToScreen(e.Location);
            _explorerLikeContextMenu.Show(screenPoint.X, screenPoint.Y, fileInfos);
        }
    }
}