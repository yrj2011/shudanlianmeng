//输入框获取&失去焦点处理
function inputFocusText(input,setting){
    var defaultVal = $(input).val(),
        defaultColor = setting && setting.defaultColor || '#aaa',
        focusColor = setting && setting.focusColor || '#333';
    $(input).css('color',defaultColor);//初始化
    $(input).click(function(){
        this.select();
    })
    $(input).blur(function(){
        if($(this).val() === '') {
            $(this).val(defaultVal).css('color',defaultColor);
        } else if($(this).val() === defaultVal) {
            $(this).css('color',defaultColor);
        } else {
            $(this).css('color',focusColor);
        }
    })
}

$(document).ready(function(){
    inputFocusText('.header-search-input',{
        defaultColor: '#aaa',
        focusColor: '#333'
    });

    //导航
    (function(){
        var $nav = $('#p-drop-nav'),
            $tri = $('#p-drop-nav-tri'),
            $menu = $('#p-drop-nav-menus'),
            timer;
        $nav.hover(function(){
            if(timer) clearTimeout(timer);
            $nav.addClass('active');
            $menu.stop(false,true).slideDown(100);
        },function(){
            timer = setTimeout(function(){
                $menu.stop(false,true).slideUp(100);
                $nav.removeClass('active');
            },300);
        });
        $('#p-drop-nav-menus li:odd').addClass('odd');
    })();



    //tabs切换
    (function(){
        var timer;
        $('.ui-tabs-tri').hover(function(){
            var self = $(this),
                i = self.index();

            timer = setTimeout(function(){
                self.addClass('ui-tabs-tri-active').siblings().removeClass('ui-tabs-tri-active');
                self.closest('.ui-tabs').find('.ui-tabs-panel').eq(i).stop(false,true).fadeIn().siblings().hide();
            },200);
        },function(){
            clearTimeout(timer);
        });



    })();

});



