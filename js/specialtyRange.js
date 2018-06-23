//-------------------------------------------------
// specialtyRange.js
//-------------------------------------------------


//-------------------------------------------------
$( document ).ready( function()
{
//    var categoryContainerWidth = $('.container.category-nav').width(); 
//    var windowWidth = $( window ).width();
//    categoryContainerWidth = ( ( windowWidth - categoryContainerWidth ) /2 ) + categoryContainerWidth; 
//    $('.container.category-nav').width( categoryContainerWidth );

    $("form#quick-search").on( "click", ".fa-search", function() 
    {
        $("form#quick-search").submit();
    });

    //-------------------------------------------------
    $(".frontpage-category-photos").on( "mouseover", "", function( event )
    {
        var left = -90;
        var top = -90;
        var offset = $(this).offset();

        var windowWidth = $( window ).width();
        var windowHeight = $( window ).height();

        if( ( offset.left - 100 ) < 0 ) { left = -25; }
        else if( (offset.left - 100 + this.offsetWidth) > windowWidth ) { left = windowWidth - this.offsetWidth - offset.left - 25; }

        if( ( offset.top - 100 ) < 0 ) { top = 0; }
        $(this).css( {"left":left+"px", "top":top+"px"});  
    });
    $(".frontpage-category-photos").on( "mouseout", "", function( event )
    {
        var e = event;
        var offset = $(this).offset();
        $(this).css( {"left":"0px", "top":"0px"});  
    });

});