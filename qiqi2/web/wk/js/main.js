/***********************************************
*READY函数

使用方法
ready(navInit);
或
ready(function () {
    navInit();  
}); 
function navInit() {
    document.getElementById("info").innerHTML = "document.getElementById(“info”).innerHTML = ok";
}
************************************************/
(function () {
    var isReady = false; //判断onDOMReady方法是否已经被执行过
    var readyList = []; //把需要执行的方法先暂存在这个数组里
    var timer; //定时器句柄
    ready = function (fn) {
        if (isReady)
            fn.call(document);
        else
            readyList.push(function () { return fn.call(this); });
        return this;
    }
    var onDOMReady = function () {
        for (var i = 0; i < readyList.length; i++) {
            readyList[i].apply(document);
        }
        readyList = null;
    }
    var bindReady = function (evt) {
        if (isReady) return;
        isReady = true;
        onDOMReady.call(window);
        if (document.removeEventListener) {
            document.removeEventListener("DOMContentLoaded", bindReady, false);
        } else if (document.attachEvent) {
            document.detachEvent("onreadystatechange", bindReady);
            if (window == window.top) {
                clearInterval(timer);
                timer = null;
            }
        }
    };

    if (document.addEventListener) {
        document.addEventListener("DOMContentLoaded", bindReady, false);
    } else if (document.attachEvent) {
        document.attachEvent("onreadystatechange", function () {
            if ((/loaded|complete/).test(document.readyState))
                bindReady();
        });
        if (window == window.top) {
            timer = setInterval(function () {
                try {
                    isReady || document.documentElement.doScroll('left');//IE中DOM未就绪时doScroll('left')会出错
                } catch (e) {
                    return;
                }
                bindReady();
            }, 5);
        }
    }
})();


ready(function(){
    $('.home-task-hot-top li').mouseenter(function(){
        $(this).addClass('on').siblings().removeClass('on');
    });

    $('.task-blocklist li').hover(function(){
        $(this).addClass('hover');
    },function(){
        $(this).removeClass('hover');
    });

    /* 列表价格范围 */
    (function(){
        var timer;
        $('.priceinput input').click(function(){
            clearTimeout(timer);
            $('.priceinterval').addClass('pintervalcur');
            $('.priceinterval span').removeClass('fn-hide');
        });
        $('.priceinterval').mouseleave(function(){
            timer = setTimeout(function(){
                $('.priceinterval').removeClass('pintervalcur');
                $('.priceinterval span').addClass('fn-hide');
            },400)
        });
        $('.priceinterval').mouseenter(function(){
            clearTimeout(timer);
        });
        $('.emptylink').click(function(){
            $('.priceinterval input').val('')
        })
    })();

    /* 列表行hover效果 */
    $('.tasklist ul').hover(function(){
        $(this).addClass('listbgc')
    },function(){
        $(this).removeClass('listbgc')
    });

    /* 更多筛选按钮 */
    $('#moreSort').live('click',function(){
        $(this).addClass('moresort-folded');
        $('#moreCase').removeClass('casebox-s');
        $('#moreSpan').removeClass('arr_top');

        $('.moresort-folded').live('click',function(){
            $(this).removeClass('moresort-folded');
            $('#moreCase').addClass('casebox-s');
            $('#moreSpan').addClass('arr_top');
        });
    });


    /* 详情顶部弹出说明 */
    $('.notebox').hover(function(){
        $(this).find('.tipboxlink').addClass('nboxtext');
        $(this).find('.de_note').removeClass('fn-hide');
    },function(){
        $(this).find('.tipboxlink').removeClass('nboxtext');
        $(this).find('.de_note').addClass('fn-hide');
    })
    

    /* 详情折叠展开 */
    $('#taskdetToggle').click(function(){
        $(this).toggleClass('toggle-off');
        $('.j_toggle-panel').slideToggle(200);
    })


    /* 帮助侧栏菜单 */
    $('.help-sub-nav dt').click(function(){
        $(this).closest('dl').addClass('on').siblings('dl').removeClass('on').find('dd').slideUp(100);
        $(this).siblings('dd').slideDown(100);
    });


    // 会员侧栏折叠
    $('.user-sub-box-hd').click(function(){
        $(this).siblings('.user-sub-box-bd').slideToggle(100).closest('.user-sub-box').toggleClass('on');
    });


});
