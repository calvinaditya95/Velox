<?php $keyword = $_GET["keyword"]; ?>
<?php $sDir = $_GET["sDir"]; ?>
<!DOCTYPE html>
<html>
<body onload="Start()">
<div id="id01"></div>
<div id="id02"></div>
<script>
    function Start() {
	<?php exec('LSE.exe "'. $keyword . '" "' . $sDir . '"'); ?>;
	Search();
    }

    function Search() {
        var xmlhttp = new XMLHttpRequest();
        var url = "SearchResult.txt";
        xmlhttp.onreadystatechange = function () {
            if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
                var myArr = JSON.parse(xmlhttp.responseText);
                DisplayResult(myArr);
            }
        }
        xmlhttp.open("GET", url, true);
        xmlhttp.send();
    }

    function DisplayResult(arr) {
        var out = "";
        var i;
        for (i = 0; i < arr.length; i++) {
            out += '<a href="Exec.php?path=' + arr[i].url + '" TARGET="_blank">' + arr[i].url + '</a><br>' + arr[i].display + '<br><br>';
        }
        document.getElementById("id01").innerHTML = out;
    }
</script>
</body>
</html>