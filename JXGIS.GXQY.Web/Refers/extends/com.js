; (function () {
    var m = window.Messager || {};

    var centerStyle = {
        right: '',
        bottom: ''
    };

    m.success = function (content) {
        $.messager.show({
            title: '成功',
            msg: content || "成功",
            timeout: 2000,
            style: centerStyle
        });
    }

    m.error = function (content) {
        $.messager.alert({
            title: '错误',
            msg: content || "错误",
            style: centerStyle
        });
    }

    m.confirm = function (content, okf) {
        $.messager.confirm('请确认', content || "确定？", function (r) {
            if (r) {
                okf();
            }
        });
    }

    Messager = m;

    Post = window.Post || function (url, obj, sf, ef) {
        $.post(url, obj, function (rt) {
            var er = rt.ErrorMessage;
            if (er) {
                if (ef) {
                    ef(er);
                }
                else {
                    m.error(er);
                }
            } else {
                sf(rt.Data);
            }
        }, 'json');
    };
})();