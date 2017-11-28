window.onload=function()
{
	var p=document.getElementById("picScroll");
	var p1=document.getElementById("picScroll1");
	var l=setInterval(changeToLeft,10);	
	function changeToLeft()
	{
		if(p.scrollLeft>=p1.offsetWidth)
		{
			p.scrollLeft=0;
		}	
		else{
			p.scrollLeft+=1;
		}	   
	}		
	p.onmouseover=function(){
		clearInterval(l);
	}
	p.onmouseout=function(){
		l=setInterval(changeToLeft,10);
	}
}


