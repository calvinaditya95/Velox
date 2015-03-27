<script>
var file1 = "tio.txt";
GetExtension(file1);

function GetExtension(fname){
	return fname.substr((~-fname.lastIndexOf(".") >>> 0) + 2);
}
</script>