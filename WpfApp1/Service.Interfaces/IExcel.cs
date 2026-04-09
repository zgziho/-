using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WpfApp1.Models;

namespace WpfApp1.Service.Interfaces
{
    /// <summary>
    /// 参数Excel数据服务接口。
    /// </summary>
    public interface IExcel
    {
        // 查询加载
        public ObservableCollection<Pars> LoadData();
        // 保存
        public void SaveData();
        // 添加
        public void AddData();
        // 删除
        public void DeleteData();
        // 更新
        public void UPdata();

    }
}
