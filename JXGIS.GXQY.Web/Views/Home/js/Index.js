function init() {
    var $condition = $('.condition');
    $('.body').height($(window).height() - $condition.height() - 20);

    var $projects = $("#cb_projects").combobox({ panelHeight: 300 });
    var $workers = $("#tb_workers");
    var $btnLoad = $('#btn_load').linkbutton({
        iconCls: 'icon-search'
    });
    var condition = { year: null, month: null, projectId: null };
    var $month = $("#cb_month").combobox({
        panelHeight: 'auto',
        onSelect: function (e) {
            condition.month = e.id;
        },
        data: (function () {
            var months = [];
            for (var i = 1; i <= 12; i++) {
                months.push({
                    id: i - 0,
                    text: i + "月",
                    value: (i < 10 ? '0' : '') + i
                })
            }
            return months;
        })()
    });

    var $year = $("#cb_year").combobox({
        panelHeight: 'auto',
        onSelect: function (e) {
            condition.year = e.id;
        },
        data: (function () {
            var months = [];
            for (var i = 2018; i <= 2020; i++) {
                months.push({
                    id: i - 0,
                    text: i + "年",
                    value: '' + i
                })
            }
            return months;
        })()
    });

    $workers.height($workers.parent().height());

    $workers.datagrid({
        singleSelect: true,
        toolbar: [
            {
                text: '添加人员',
                iconCls: 'icon-add',
                handler: function () {
                    WorkerList.show();
                }
            }, {
                text: '移除人员',
                iconCls: 'icon-remove',
                handler: function () { alert('cut') }
            }, {
                text: '刷新表格',
                iconCls: 'icon-reload',
                handler: function () { alert('cut') }
            }, ],
        frozenColumns: [
            [
                { field: 'Name', title: '姓名', width: 120, align: 'center' },
                { field: 'BasePay', title: '基本工资', width: 120, align: 'center' },
                { field: 'Bonus', title: '绩效奖金', width: 120, align: 'center' }
            ]
        ],
        columns: [[
            {
                field: 'WorkDays', title: '考勤情况', width: 1500, align: 'center', formatter: function (value, row, index) {
                    var html = "<div class='workerdays'>";
                    for (var i = 1; i < 31; i++) {
                        html += "<div data-index='" + index + "' data-row='" + i + "'>" + i + "</div>";
                    }
                    html += "</div>";
                    return html;
                }
            },
        ]],
        data: (function () {
            var arr = [];
            for (var i = 0; i < 100; i++) {
                arr.push({ id: i, name: i });
            }
            return arr;
        })(),
        onLoadSuccess: function () {
            $('.workerdays>div').on('click', function () {
                var $this = $(this);
                console.log($this.data('index'), $this.data('row'));
            })
        }
    });

    $projects.combobox({
        onSelect: function (e) {
            condition.projectId = e.id;
        },
        onChange: function (n, o) {
            //console.log(e);
        },
    });

    function getProjects() {
        Project.getProjects(function (d) {
            $projects.combobox("loadData", d ? d.map(function (i) {
                return {
                    id: i.Id,
                    text: i.Name,
                    value: i.Id
                }
            }) : [])
        });
    };

    $('#p_btn_add').on('click', function () {
        Project.show();
    });

    $('#p_btn_remove').on('click', function () {
        var id = condition.projectId;
        if (id) {
            Messager.confirm("确定删除当前项目？", function () {
                Project.removeProject(id, function () {
                    Messager.success("已删除项目！");
                    getProjects();
                    $projects.combobox('clear');
                    condition.projectId = null;
                });
            });
        } else {
            Messager.error("尚未选择任何项目！");
        }
    });

    $('#p_btn_edit').on('click', function () {
        var id = condition.projectId;
        if (id) {
            Project.show(id);
        } else {
            Messager.error("尚未选择任何项目！");
        }
    });


    $(Project).on("saveSuccess", function () {
        getProjects();
    });
    getProjects();
}

$(init);