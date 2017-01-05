$(document).ready(function () {
    // 初始化
    init();
});

function init() {
    // 將集團放入下拉式選單
    GetClub();
}

// 取得集團資訊
function GetClub() {

    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "http://192.168.51.21:63000/main/TasWebService.aspx/GetAllCompanyGroup",
        data: {},
        dataType: "json",
        success: function (info) {
            console.log(info.d);
            FillClub(info.d);
        },
        error: function (xhr, ajaxOptions, thrownError) {
            console.log('error');
        }
    });
}

// 填入集團資訊
function FillClub(data) {
    $("#club").empty();
    $.each(data, function (index, value) {
        var str = "";
        if (index == 0)
            str += "<option  value='0'>選擇類型</option>"
        str += "<option value='" + value.id + "'>" + value.name + "</option>";
        $('#club').append(str);
    });
}

// 改變時觸發,取得有填完問卷內容的問卷
function ChangeClub() {
    var clubId = $('#club').val();
    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "http://192.168.51.21:63000/main/TasWebService.aspx/GetProjectListByCompanyGroup",
        data: '{"companyGroupID":' + clubId + '}',
        dataType: "json",
        success: function (info) {
            console.log(info);

            FillTable(info.d);
        },
        error: function (xhr, ajaxOptions, thrownError) {
            console.log('error');
        }
    });
}

// 將資料放入到Table
function FillTable(data) {

    $("#myTable tbody").html("");

    // Add some examples to the page
    $.each(data, function (index, value) {

        var $tr = $('<tr>');
        var $tdId = $('<td>').text(value.id);
        var $tdProjectName = $('<td>').text(value.projectName);
        var $tdProjecTypeText = $('<td>').text(value.projecTypeText);
        var $tdGroupName = $('<td>').text(value.groupName);
        var $tdStaffCount = $('<td>').text(value.staffCount);
        $tr.append($tdId)
        $tr.append($tdProjectName)
        $tr.append($tdProjecTypeText)
        $tr.append($tdGroupName);
        $tr.append($tdStaffCount);
        var $td = $('<td>');
        var $inputDownload = $('<input>').attr('type', 'button').addClass('faq_btn select_btn').val('成績單')

        // 綁定事件
        $inputDownload.click(function () {
            GetQuestionResult(value.groupId);
        })

        $inputDownload.appendTo($td)
        $tr.append($td);

        $('#myTable > tbody').append($tr);
    });

    $('#myTable').paginate({
        // 設定一次顯示幾筆
        limit: 5,
        onSelect: function (obj, page) {
        }
    });

}

// 取得問卷結果
function GetQuestionResult(groupId) {
    $.ajax({
        type: "POST",
        contentType: "application/json; charset=utf-8",
        url: "http://192.168.51.21:63000/main/TasWebService.aspx/GetMatualEvaluateDetail",
        data: '{"groupID":' + groupId + '}',
        dataType: "json",
        success: function (info) {
            console.log(info);
            CreateTable(info.d);
        },
        error: function (xhr, ajaxOptions, thrownError) {
            console.log('error');
        }
    });
}

// 產生表格內容
function CreateTable(data) {
    if (data.length == 0) {
        alert("程式發生問題請聯絡開發人員");
        return "";
    }
    var now = false;
    // 清空表格內容
    $(".table2excel").html("");

    var str = "";
    $.each(data, function (arrIndex, arr) {

        str += '<tr>'
        $.each(arr, function (index, val) {
            if (now && arrIndex == 1 && index == (arr.length - 1)) {
                str += '<td align="center" rowspan="' + (data.length - 1) + '">' + val + '</td>'
            } else
                str += '<td align="center">' + val + '</td>'
        });
        str += '</tr>'
    });
    $(".table2excel").append(str);

    downloadExcel();
}

// 下載Excel
function downloadExcel() {
    $(".table2excel").table2excel({
        exclude: ".noExl",
        name: "Excel Document Name",
        filename: "myFileName",
        fileext: ".xls",
        exclude_img: true,
        exclude_links: true,
        exclude_inputs: true
    });
}