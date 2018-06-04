<?php
if (isset ($_POST['submit'])){
	$file=$_FILES["file"];
	
	$fileName= $_FILES["file"]["name"];
	
	$fileTmpName= $_FILES["file"]["tmp_name"];
	$fileSize= $_FILES["file"]["size"];
	$fileError= $_FILES["file"]["error"];
	$fileType= $_FILES["file"]["type"];
	
	$fileExt=explode('.',$fileName);
	$fileActualExt=strtolower(end($fileExt));
	
	$allowed = array('xls');
	
	if (in_array($fileActualExt,$allowed) ){
		if($fileError===0){
			$fileNameNew = 'dimtable'.'.'.$fileActualExt;
			$fileDestination = "uploads/".$fileNameNew;
			move_uploaded_file($fileTmpName,$fileDestination);
			
		}else{
			echo "there was an error";
		}
		
	}else{
		echo "cannot upload this type of file";
		
	}
	

}
?>

<html>
<head>
<title> jquery</title>



<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>


<link href="jquery-ui/jquery-ui.css" rel="stylesheet">
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<link rel="stylesheet" type="text/css" href="karsunstyle.css">
<style type="text/css">

#forms{
	position:relative;
	width:25%;
	height:25%;
	top:300px;
	border:1px solid #8e8d8d;
	border-radius:10px;
	background-color:#D6D6D6;
	margin: 0px auto;
	padding:15px;
	font-size:14px;
	font-weight:bold;
	text-align:center;
}

input{
	position:relative;
	
	padding:15px;
	margin-top:20%;
	border-radius:10px;
	background-color:white;
	width:50%;
	border: transparent !important;
	
}
input:hover{
	
	background-color:#93b4ed;
	color:white;
	
	
}
#in1{
	text-align:center;
	
	
}

 

</style>





</head>



<body oncontextmenu="return false">
<div id="fixed">
<div id="top"><img id="top_logo" height="100px" src="http://www.karsun-llc.com/wp-content/uploads/2016/09/Logo-300-150.png"></div>

<div id="second_bar">

<div id ="home" class="second_bar_section">
<div class="same">
<a href="home.html">Home </a></div>

<div class="same">
<a id="hover1" href="display.html">Display Board </a>

<div class="dropdown"  >
<p><a href="display.html" style="color:black">Interactive Board </a></p>
<p><a href="piechart.html" style="color:black">Pie-Chart </a></p>
<p><a href="bargraph.html" style="color:black">Bar-graph </a></p>
</div>

</div>


<div class="same">
<a href="uploads.html">Uploads </a></div>

<div class="same">
<a href="">Request </a></div>

<div class="same">
<a href="dimtable.html">Create Dim Table</a></div>
</div>



</div>

</div>



<div id="forms">

<form action="./cgi-bin/dimtablecreation.py">
<p>please click to create a dim table with the data in the file you uploaded<span id ="filename"></span></p>
<div id="in1">
<input type="submit" value="Create table">
</div>
</form>
</div>





<div id="bottomfixed">
<div id="bottombar"><span id="bottomtext">KARSUN SOLUTIONS-llc &copy; </span></div>
</div>
<script type="text/javascript" src="karsunscripts.js"></script>
<script type="text/javascript">


  
</script>


</body>






</html>
