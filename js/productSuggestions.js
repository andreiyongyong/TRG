//-------------------------------------------------
// productSuggestions.js
//-------------------------------------------------

//var idxProductSuggestionSideBar = 0;
var idxProductSuggestionPageBottom = 0;

var imageIconWidth = 150;
var maxDisplayWidth = 450;

var PAGE_BOTTOM_SHIFT_DISTANCE = "180px";

//setInterval(function(){ shiftSuggestionsLeftSideBar(); }, 2000);
setInterval(function(){ shiftSuggestionsLeftPageBottom(); }, 8000);

//-------------------------------------------------
$( document ).ready( function()
{ 
    // add initial 3 items to list
    for( i=0; i < 3; i++)
    {   
        if( i < AllSuggestedProducts.length )
        {
            //$("#productSuggestionsSideBarInnerDiv").append( AllSuggestedProducts[i] );
            $("#productSuggestionsPageBottomInnerDiv").append( AllSuggestedProducts[i] );
            idxProductSuggestionPageBottom++;
        }
    }

    $("i.arrow_left").on("click", "", function() 
    {
        shiftSuggestionsRightPageBottom();
    });

    //-------------------------------------------------
    $("i.arrow_right").on("click", "", function() 
    {
        shiftSuggestionsLeftPageBottom();
    });


}); // $( document ).ready( function()

//-------------------------------------------------
function shiftSuggestionsLeftPageBottom()
{
    if( idxProductSuggestionPageBottom >= AllSuggestedProducts.length ) { idxProductSuggestionPageBottom = 0 }

    $("#productSuggestionsPageBottomInnerDiv").append( AllSuggestedProducts[idxProductSuggestionPageBottom] ); 
    $("#productSuggestionsPageBottomInnerDiv").animate( { left : "-="+PAGE_BOTTOM_SHIFT_DISTANCE }, 2000, function() {
        $("#productSuggestionsPageBottomInnerDiv").css({ left: "+="+PAGE_BOTTOM_SHIFT_DISTANCE });
        $("#productSuggestionsPageBottomInnerDiv > div.productSuggestions_productWrapper:first-child").remove(); 
    } );
    idxProductSuggestionPageBottom++; 

} // shiftSuggestionsLeftPageBottom

//-------------------------------------------------
function shiftSuggestionsRightPageBottom()
{
    var posLeft = parseInt( $("#productSuggestionPageBottomInnerDiv").css("left"), 10 );
    if( posLeft < 0 )   
    {
        posLeft += imageIconWidth;
        $("#productSuggestionPageBottomInnerDiv").css("left", posLeft )
        $("i.arrow_left").on('mouseenter mouseleave');
    }
    else
    {
        $("i.arrow_left").off('mouseenter mouseleave');
    }

} // shiftSuggestionsRightPageBottom

//-------------------------------------------------
function shiftSuggestionsLeftSideBar()
{
    if( idxProductSuggestionSideBar >= AllSuggestedProducts.length ) { idxProductSuggestionSideBar = 0 }

    $("#productSuggestionsSideBarInnerDiv").append( AllSuggestedProducts[idxProductSuggestionSideBar] ); 
    $("#productSuggestionsSideBarInnerDiv a:first-child").remove(); 

    idxProductSuggestionSideBar++; 

} // shiftSuggestionsLeftSideBar

//-------------------------------------------------
function shiftSuggestionsRightSideBar()
{

} // shiftSuggestionsRightSideBar
                          