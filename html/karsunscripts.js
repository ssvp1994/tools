$("#hover1").hover(function(){
    $(".dropdown").finish().slideDown('fast');
 
 });
 
  $(".dropdown").hover(function(){
  
  
    $(".dropdown").css("display","block");
	}, function() {
     $(".dropdown").finish().slideUp('fast');
  });