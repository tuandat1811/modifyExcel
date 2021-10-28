using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ModifyExcel
{
    /// <summary>
    ///     Cấu trúc thông tin của một sheet
    /// </summary>
    [JsonObject]
    class SheetType
    {
        /// <summary> Tên của sheet. Vi dụ "STEP 4". </summary>
        public string name;

        /// <summary> Danh sách các cell và nội dung </summary>
        /// <see cref="Newtonsoft.Json.JsonSerializationException"> WorkbookData.SheetType[].CellType[]  thì SheetType không cần hàm khởi tạo, nhưng CellType lại băt buộc phải có</see>
        public List<CellType> cells;

        /// <summary> Lỗi khi chuyển từ ORM vào excel </summary>
        public string errMessage = null;
        
        public SheetType(List<CellType> cells, string name = null)
        {
            this.name = name;
            this.cells = cells;
            errMessage = null;
        }
    }
}
