using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RepairServiceTracker;

internal static class Program
{
    [STAThread] static void Main() { ApplicationConfiguration.Initialize(); Application.Run(new MainForm()); }
}

public class RepairOrder
{
    public string MaDon { get; set; } = "";
    public string TenKhachHang { get; set; } = "";
    public string SoDienThoai { get; set; } = "";
    public string LoaiThietBi { get; set; } = "";
    public string MoTaLoi { get; set; } = "";
    public DateTime NgayTiepNhan { get; set; }
    public decimal ChiPhiDuKien { get; set; }
    public string TrangThai { get; set; } = "Mới tiếp nhận";
    public string DuongDanAnh { get; set; } = "";
}

public class MainForm : Form
{
    readonly BindingList<RepairOrder> data = new();
    readonly BindingSource source = new();
    readonly ErrorProvider errors = new();
    readonly TextBox txtCode = new(), txtCustomer = new(), txtPhone = new(), txtDescription = new(), txtSearch = new(), txtImage = new();
    readonly ComboBox cbType = new(), cbStatus = new(), cbFilter = new();
    readonly DateTimePicker dtReceived = new();
    readonly NumericUpDown numCost = new() { Minimum = 0, Maximum = 1000000000, DecimalPlaces = 0, ThousandsSeparator = true, Increment = 10000 };
    readonly PictureBox picture = new();
    readonly DataGridView grid = new();
    readonly ToolStripStatusLabel statusLabel = new();
    string currentImage = "";

    public MainForm()
    {
        Text="Repair Service Tracker"; StartPosition=FormStartPosition.CenterScreen; Size=new Size(1280,760); MinimumSize=new Size(950,620);
        Font=new Font("Segoe UI",9F); errors.ContainerControl=this;
        cbType.DropDownStyle=ComboBoxStyle.DropDownList; cbType.Items.AddRange(new object[]{"Laptop","Điện thoại","Máy tính bảng","Máy tính bàn","Thiết bị khác"}); cbType.SelectedIndex=0;
        cbStatus.DropDownStyle=ComboBoxStyle.DropDownList; cbStatus.Items.AddRange(new object[]{"Mới tiếp nhận","Đang sửa","Chờ linh kiện","Hoàn tất","Đã hủy"}); cbStatus.SelectedIndex=0;
        cbFilter.DropDownStyle=ComboBoxStyle.DropDownList; cbFilter.Items.AddRange(new object[]{"Tất cả trạng thái","Mới tiếp nhận","Đang sửa","Chờ linh kiện","Hoàn tất","Đã hủy"}); cbFilter.SelectedIndex=0;
        dtReceived.Format=DateTimePickerFormat.Short; dtReceived.Value=DateTime.Today;
        BuildUi(); source.DataSource=data; grid.DataSource=source; AddSamples(); RefreshView();
    }

    void BuildUi()
    {
        var menu=new MenuStrip(); var file=new ToolStripMenuItem("Tệp"); var export=new ToolStripMenuItem("Xuất CSV"); export.Click+=(_,_)=>ExportCsv();
        var exit=new ToolStripMenuItem("Thoát"); exit.Click+=(_,_)=>Close(); file.DropDownItems.AddRange(new ToolStripItem[]{export,exit}); menu.Items.Add(file); MainMenuStrip=menu;
        var status=new StatusStrip(); status.Items.Add(statusLabel);
        var root=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2,Padding=new Padding(10)}; root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent,100));
        var entry=new TableLayoutPanel{Dock=DockStyle.Top,AutoSize=true,ColumnCount=4,RowCount=5,Padding=new Padding(8),CellBorderStyle=TableLayoutPanelCellBorderStyle.Single};
        for(int i=0;i<4;i++)entry.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,25));
        Field(entry,"Mã đơn",txtCode,0,0); Field(entry,"Tên khách hàng",txtCustomer,1,0); Field(entry,"Số điện thoại",txtPhone,2,0); Field(entry,"Loại thiết bị",cbType,3,0);
        Field(entry,"Ngày tiếp nhận",dtReceived,0,2); Field(entry,"Chi phí dự kiến (VNĐ)",numCost,1,2); Field(entry,"Trạng thái",cbStatus,2,2);
        var descPanel=new Panel{Dock=DockStyle.Fill,Height=75,Padding=new Padding(5)}; descPanel.Controls.Add(new Label{Text="Mô tả lỗi",Dock=DockStyle.Top,Height=22}); txtDescription.Multiline=true; txtDescription.ScrollBars=ScrollBars.Vertical; txtDescription.Dock=DockStyle.Fill; descPanel.Controls.Add(txtDescription); entry.Controls.Add(descPanel,3,2);
        var imagePanel=new FlowLayoutPanel{Dock=DockStyle.Fill,AutoSize=true,Padding=new Padding(5)};
        txtImage.Width=240; txtImage.ReadOnly=true; txtImage.PlaceholderText="Chưa chọn ảnh";
        var choose= new Button{Text="Chọn ảnh...",Width=100,Height=28}; choose.Click+=(_,_)=>ChooseImage();
        picture.Size=new Size(110,85); picture.BorderStyle=BorderStyle.FixedSingle; picture.SizeMode=PictureBoxSizeMode.Zoom;
        imagePanel.Controls.Add(new Label{Text="Ảnh thiết bị:",AutoSize=true,Padding=new Padding(0,7,0,0)}); imagePanel.Controls.Add(txtImage); imagePanel.Controls.Add(choose); imagePanel.Controls.Add(picture);
        entry.Controls.Add(imagePanel,0,4); entry.SetColumnSpan(imagePanel,4);
        var buttons=new FlowLayoutPanel{Dock=DockStyle.Fill,AutoSize=true};
        buttons.Controls.AddRange(new Control[]{Btn("Thêm",(_,_)=>AddItem()),Btn("Cập nhật",(_,_)=>UpdateItem()),Btn("Xóa",(_,_)=>DeleteItem()),Btn("Làm mới",(_,_)=>Clear())});
        entry.Controls.Add(buttons,0,5); entry.SetColumnSpan(buttons,4); entry.RowCount=6; entry.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        var list=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2}; list.RowStyles.Add(new RowStyle(SizeType.AutoSize)); list.RowStyles.Add(new RowStyle(SizeType.Percent,100));
        var filters=new FlowLayoutPanel{Dock=DockStyle.Fill,AutoSize=true,Padding=new Padding(0,8,0,8)};
        filters.Controls.Add(new Label{Text="Tìm mã đơn / tên / SĐT:",AutoSize=true,Padding=new Padding(0,7,0,0)});
        txtSearch.Width=260; txtSearch.PlaceholderText="Nhập từ khóa..."; txtSearch.TextChanged+=(_,_)=>RefreshView(); filters.Controls.Add(txtSearch);
        filters.Controls.Add(new Label{Text="Trạng thái:",AutoSize=true,Padding=new Padding(0,7,0,0)});
        cbFilter.SelectedIndexChanged+=(_,_)=>RefreshView(); filters.Controls.Add(cbFilter); list.Controls.Add(filters,0,0);
        ConfigureGrid(); grid.SelectionChanged+=(_,_)=>LoadSelected(); list.Controls.Add(grid,0,1);
        root.Controls.Add(entry,0,0); root.Controls.Add(list,0,1); Controls.Add(root); Controls.Add(status); Controls.Add(menu); status.Dock=DockStyle.Bottom;
    }

    void Field(TableLayoutPanel p,string label,Control c,int col,int row)
    {
        var panel=new Panel{Dock=DockStyle.Fill,Height=58,Padding=new Padding(5)};
        panel.Controls.Add(new Label{Text=label,Dock=DockStyle.Top,Height=22}); c.Dock=DockStyle.Top; panel.Controls.Add(c); p.Controls.Add(panel,col,row);
    }
    Button Btn(string t,EventHandler e){var b=new Button{Text=t,Width=105,Height=32,Margin=new Padding(5)};b.Click+=e;return b;}
    void ConfigureGrid()
    {
        grid.Dock=DockStyle.Fill;grid.AutoGenerateColumns=false;grid.SelectionMode=DataGridViewSelectionMode.FullRowSelect;grid.MultiSelect=false;grid.ReadOnly=true;grid.AllowUserToAddRows=false;grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;
        Col("Mã đơn",nameof(RepairOrder.MaDon));Col("Khách hàng",nameof(RepairOrder.TenKhachHang));Col("Số điện thoại",nameof(RepairOrder.SoDienThoai));
        Col("Thiết bị",nameof(RepairOrder.LoaiThietBi));Col("Mô tả lỗi",nameof(RepairOrder.MoTaLoi));Col("Ngày nhận",nameof(RepairOrder.NgayTiepNhan),"dd/MM/yyyy");
        Col("Chi phí (VNĐ)",nameof(RepairOrder.ChiPhiDuKien),"N0");Col("Trạng thái",nameof(RepairOrder.TrangThai));
    }
    void Col(string h,string p,string f="")=>grid.Columns.Add(new DataGridViewTextBoxColumn{HeaderText=h,DataPropertyName=p,DefaultCellStyle={Format=f}});
    void ChooseImage()
    {
        using var d=new OpenFileDialog{Title="Chọn ảnh thiết bị",Filter="Tệp ảnh|*.jpg;*.jpeg;*.png;*.bmp;*.gif"};
        if(d.ShowDialog()!=DialogResult.OK)return;
        currentImage=d.FileName;txtImage.Text=currentImage;
        try{using var fs=new FileStream(currentImage,FileMode.Open,FileAccess.Read);using var img=Image.FromStream(fs);picture.Image=new Bitmap(img);}
        catch{MessageBox.Show("Không thể mở ảnh đã chọn.");}
    }
    RepairOrder Read()=>new(){MaDon=txtCode.Text.Trim(),TenKhachHang=txtCustomer.Text.Trim(),SoDienThoai=txtPhone.Text.Trim(),LoaiThietBi=cbType.Text,
        MoTaLoi=txtDescription.Text.Trim(),NgayTiepNhan=dtReceived.Value.Date,ChiPhiDuKien=numCost.Value,TrangThai=cbStatus.Text,DuongDanAnh=currentImage};
    bool Valid()
    {
        errors.Clear();bool ok=true;
        if(string.IsNullOrWhiteSpace(txtCode.Text)){errors.SetError(txtCode,"Nhập mã đơn.");ok=false;}
        if(string.IsNullOrWhiteSpace(txtCustomer.Text)){errors.SetError(txtCustomer,"Nhập tên khách hàng.");ok=false;}
        if(!System.Text.RegularExpressions.Regex.IsMatch(txtPhone.Text.Trim(),@"^(0|\+84)[0-9]{9,10}$")){errors.SetError(txtPhone,"Số điện thoại không hợp lệ.");ok=false;}
        if(string.IsNullOrWhiteSpace(txtDescription.Text)){errors.SetError(txtDescription,"Nhập mô tả lỗi.");ok=false;}
        if(numCost.Value<0){errors.SetError(numCost,"Chi phí không được âm.");ok=false;}
        return ok;
    }
    void AddItem()
    {
        if(!Valid())return;
        if(data.Any(x=>x.MaDon.Equals(txtCode.Text.Trim(),StringComparison.OrdinalIgnoreCase))){errors.SetError(txtCode,"Mã đơn đã tồn tại.");return;}
        data.Add(Read());RefreshView();Clear();
    }
    void UpdateItem()
    {
        if(grid.CurrentRow?.DataBoundItem is not RepairOrder old||!Valid())return;
        if(data.Any(x=>x!=old&&x.MaDon.Equals(txtCode.Text.Trim(),StringComparison.OrdinalIgnoreCase))){errors.SetError(txtCode,"Mã đơn đã tồn tại.");return;}
        var n=Read();old.MaDon=n.MaDon;old.TenKhachHang=n.TenKhachHang;old.SoDienThoai=n.SoDienThoai;old.LoaiThietBi=n.LoaiThietBi;old.MoTaLoi=n.MoTaLoi;old.NgayTiepNhan=n.NgayTiepNhan;old.ChiPhiDuKien=n.ChiPhiDuKien;old.TrangThai=n.TrangThai;old.DuongDanAnh=n.DuongDanAnh;
        source.ResetBindings(false);RefreshView();
    }
    void DeleteItem()
    {
        if(grid.CurrentRow?.DataBoundItem is not RepairOrder x)return;
        if(MessageBox.Show("Bạn có chắc muốn xóa đơn này?","Xác nhận",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes){data.Remove(x);RefreshView();Clear();}
    }
    void RefreshView()
    {
        var q=Filtered().ToList();source.DataSource=new BindingList<RepairOrder>(q);grid.DataSource=source;
        decimal total=data.Where(x=>x.TrangThai!="Hoàn tất"&&x.TrangThai!="Đã hủy").Sum(x=>x.ChiPhiDuKien);
        statusLabel.Text=$"Số đơn sau lọc: {q.Count} | Tổng chi phí dự kiến (chưa hoàn tất): {total:N0} VNĐ";
    }
    System.Collections.Generic.IEnumerable<RepairOrder> Filtered()
    {
        var q=data.AsEnumerable();string key=txtSearch.Text.Trim();
        if(key.Length>0)q=q.Where(x=>x.MaDon.Contains(key,StringComparison.OrdinalIgnoreCase)||x.TenKhachHang.Contains(key,StringComparison.OrdinalIgnoreCase)||x.SoDienThoai.Contains(key));
        if(cbFilter.SelectedIndex>0)q=q.Where(x=>x.TrangThai==cbFilter.Text);
        return q;
    }
    void LoadSelected()
    {
        if(grid.CurrentRow?.DataBoundItem is not RepairOrder x)return;
        txtCode.Text=x.MaDon;txtCustomer.Text=x.TenKhachHang;txtPhone.Text=x.SoDienThoai;cbType.Text=x.LoaiThietBi;txtDescription.Text=x.MoTaLoi;dtReceived.Value=x.NgayTiepNhan;numCost.Value=Math.Min(numCost.Maximum,Math.Max(numCost.Minimum,x.ChiPhiDuKien));cbStatus.Text=x.TrangThai;
        currentImage=x.DuongDanAnh;txtImage.Text=currentImage;
        picture.Image?.Dispose();picture.Image=null;
        if(File.Exists(currentImage)){try{using var fs=new FileStream(currentImage,FileMode.Open,FileAccess.Read);using var img=Image.FromStream(fs);picture.Image=new Bitmap(img);}catch{}}
    }
    void Clear()
    {
        errors.Clear();txtCode.Clear();txtCustomer.Clear();txtPhone.Clear();txtDescription.Clear();txtImage.Clear();currentImage="";picture.Image?.Dispose();picture.Image=null;
        cbType.SelectedIndex=0;cbStatus.SelectedIndex=0;dtReceived.Value=DateTime.Today;numCost.Value=0;txtCode.Focus();
    }
    void ExportCsv()
    {
        using var d=new SaveFileDialog{Filter="CSV (*.csv)|*.csv",FileName="DonSuaChua.csv"};if(d.ShowDialog()!=DialogResult.OK)return;
        var sb=new StringBuilder("MaDon,TenKhachHang,SoDienThoai,LoaiThietBi,MoTaLoi,NgayTiepNhan,ChiPhiDuKien,TrangThai,DuongDanAnh\r\n");
        foreach(var x in Filtered())sb.AppendLine(string.Join(",",new[]{x.MaDon,x.TenKhachHang,x.SoDienThoai,x.LoaiThietBi,x.MoTaLoi,x.NgayTiepNhan.ToString("dd/MM/yyyy"),x.ChiPhiDuKien.ToString("N0"),x.TrangThai,x.DuongDanAnh}.Select(Csv)));
        File.WriteAllText(d.FileName,sb.ToString(),new UTF8Encoding(true));MessageBox.Show("Đã xuất CSV.");
    }
    string Csv(string s)=>"\""+(s??"").Replace("\"","\"\"")+"\"";
    void AddSamples()
    {
        data.Add(new RepairOrder{MaDon="SC001",TenKhachHang="Nguyễn Minh Anh",SoDienThoai="0912345678",LoaiThietBi="Laptop",MoTaLoi="Máy không khởi động",NgayTiepNhan=DateTime.Today.AddDays(-1),ChiPhiDuKien=850000,TrangThai="Đang sửa"});
        data.Add(new RepairOrder{MaDon="SC002",TenKhachHang="Trần Hoàng Nam",SoDienThoai="0987654321",LoaiThietBi="Điện thoại",MoTaLoi="Màn hình không hiển thị",NgayTiepNhan=DateTime.Today,ChiPhiDuKien=1200000,TrangThai="Mới tiếp nhận"});
    }
}
