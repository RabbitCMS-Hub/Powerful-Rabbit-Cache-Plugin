function myFunction() {
    var input, filter, table, tr, td, i, txtValue;
    input = document.getElementById("myInput");
    filter = input.value.toUpperCase();
    table = document.getElementById("myTable");
    tr = table.getElementsByTagName("tr");

    for (i = 0; i < tr.length; i++) {
        td = tr[i].getElementsByTagName("td")[0];
        if (td) {
            txtValue = td.textContent || td.innerText;
            if (txtValue.toUpperCase().indexOf(filter) > -1) {
                tr[i].style.display = "";
            } else {
                tr[i].style.display = "none";
            }
        }
    }
};

$('.deleteStaticFile').unbind('click');
$('.deleteStaticFile').click(function(event) {
    let RowID = $(this).attr('data-row');
    let FileName = $(this).attr('data-filename');

    $.ajax({
        url : "ajax.asp?Cmd=PluginSettings&PluginName="+CACHE_PLUGIN_CODE+"_&Page=DELETE:OneFile",
        type: "POST",
        data: {
            fileName: FileName
        },
        success: function(data) {
            alert('Başarılı');
            $('#' + RowID).remove();
        },
        error: function(e) {
            alert('Hata');
        }
    });
});