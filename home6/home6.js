var con=false;
var arr=['很不满意','不满意','一般','满意','非常满意'];
function starPic(which){
	if(con==true) return 0;
	var x=document.getElementsByTagName("img");
	for (var i = 0; x[i-1] != which; i++) {
		x[i].src="images/star2.png";
	}
	if(i-1<=1){
		for(var j=0;j<=i-1;j++){
			x[j].src="images/star1.png";
		}
	}
	document.getElementById("text").value=arr[i-1];
}
function hide(){
	if(con==true) return 0;
	var x=document.getElementsByTagName("img");
	for (var i = x.length - 2; i >= 0; i--) {
		x[i].src="images/star0.png";
	}
	document.getElementById("text").value="";
}
function qd(which){
	con=true;
	var x=document.getElementsByTagName("img");
	for (var i = 0; x[i-1] != which; i++) {
		x[i].src="images/star2.png";
	}
	if(i-1<=1){
		for(var j=0;j<=i-1;j++){
			x[j].src="images/star1.png";
		}
	}
	document.getElementById("text").value=arr[i-1];
}