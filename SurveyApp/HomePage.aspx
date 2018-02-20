<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="HomePage.aspx.cs" Inherits="SurveyApp.HomePage" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
	Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">

	<link href="Styles/axure_rp_page.css" rel="stylesheet" type="text/css" />
	<link href="Styles/axurerp_privacy_notice.css" rel="stylesheet" type="text/css" />
	<link href="Styles/jquery-ui-themes.css" rel="stylesheet" type="text/css" />
	<link href="Styles/StyleSheet1.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="Scripts/jquery-1.4.1.js"></script>
	<script src="Scripts/jquery-1.4.1.min.js" type="text/javascript"></script>
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>

	<link rel="stylesheet" type="text/css" href="/css/print.css" media="print" />
	<style>
		.lblposition {
			float: left;
			padding: 5px 0;
			width: 475px;
		}

		.lblposition1 {
			float: left;
			padding: 5px 0;
			width: 475px;
		}


		.txtposition {
			float: left;
			padding: 5px 0;
		}

		.validation-error {
			color: Red;
			font-weight: bold;
		}

		.clear {
			clear: both;
		}
	</style>


	<%--    <script type="text/javascript">


        document.body.addEventListener('contextmenu', function (e) {

            e.preventDefault();

        });
    </script>--%>


	<%--    <style type="text/css">
@media print{
body {display:none;}
}
</style>--%>

	<style type="text/css" media="print">
		BODY {
			display: none;
			visibility: hidden;
		}
	</style>


	<style type="text/css" media="print">
		@page {
			size: auto; /* auto is the initial value */
			margin: 0; /* this affects the margin in the printer settings */
		}
	</style>

	<script type="text/javascript">


		function validate(key, obj) {
			//alert($(obj).attr('id'));
			//getting key code of pressed key


			var keycode = (key.which) ? key.which : key.keyCode;

			var phn = document.getElementById('Q44');
			// alert();
			var classstyle = $(obj).attr('class');
			// alert(classstyle);



			if (!(classstyle == "Q37" || classstyle == "Q38" || classstyle == "Q47" || classstyle == "Q48")) {
				if (classstyle == "Q49") {
					$(obj).attr('maxlength', '10');
				}
				//else if(classstyle == "Q26_1" || classstyle =="Q25_1" || classstyle == "Q27_1")
				//{

				//    $(obj).attr('maxlength', '10');
				//}

				//else
				//{
				//    $(obj).attr('maxlength', '6');

				//}
				if (!(keycode == 8 || keycode == 46) && (keycode < 48 || keycode > 57)) {


					return false;
				}
				else {


					//Condition to check textbox contains ten numbers or not
					//    if (phn.value.length < 10) {

					//    return true;
					//}
					//else {
					//    return false;
					//}
				}
			}

			else {
				$(obj).attr('maxlength', '200');


				return true;
			}


		}

	</script>

	<script type="text/javascript">


		function validate1(key, obj) {

			var keycode = (key.which) ? key.which : key.keyCode;
			var phn = document.getElementById('txtsubquestion');

			var classstyle = $(obj).attr('class');

			if (!(classstyle == "Q48_1" || classstyle == "Q48_2" || classstyle == "Q48_3" || classstyle == "Q48_4")) {
				//if (classstyle == "Q40_1" || classstyle == "Q40_2" || classstyle == "Q40_3" || classstyle == "Q40_4" || classstyle == "Q79_1"
				//    || classstyle == "Q79_2" || classstyle == "Q79_3" || classstyle == "Q88_1" || classstyle == "Q88_2" || classstyle == "Q88_3"
				//    || classstyle == "Q88_4" || classstyle == "Q51_1" || classstyle == "Q51_2" || classstyle == "Q51_3" || classstyle == "Q51_4"
				//    || classstyle == "Q51_5" || classstyle == "Q51_6" || classstyle == "Q51_7" || classstyle == "Q51_8" || classstyle == "Q51_9")
				//{

				//    $(obj).attr('maxlength', '10');
				//}

				//else {
				//    $(obj).attr('maxlength', '6');

				//}
				if (!(keycode == 8 || keycode == 46) && (keycode < 48 || keycode > 57)) {
					return false;
				}
				else {
					//Condition to check textbox contains ten numbers or not
					//if (phn.value.length < 10) {

					//    return true;
					//}
					//else {
					//    return false;
					//}
				}
			}

			else {
				$(obj).attr('maxlength', '100');

			}
		}



	</script>

	<script type="text/javascript">


		$(document).ready(function () {
			// alert('test')
			document.getElementById("parenterrorlbl").innerHTML = "";
			// document.getElementById("lblsaveerror").innerHTML = "";

			document.getElementById("parenterrorlbl").style.display = "none";


			// document.getElementById("lblsaveerror").style.display = "none";

			$("#btnNext").click(function (e) {
				document.getElementById("parenterrorlbl").innerHTML = "";
				document.getElementById("lblsaveerror").innerHTML = "";
				document.getElementById("parenterrorlbl").style.display = "none";
				document.getElementById("lblsaveerror").style.display = "none";


				var error = new Array();

				var data = $("input:hidden[id^='hdq']");
				$.each(data, function (key, object) {
					// alert(object.value);
					//alert(object.id);
					var completestring = object.id;
					//alert(completestring.split('/').length);
					var data1 = completestring.split('/')
					if (object.value == "ChkQuesRef1" || object.value == "ChQuesRef2") {

						// alert(object.value)
						var refcontrolid = data1["1"];
						var c = document.getElementsByClassName(refcontrolid)[0].value;
						c = parseInt(c);
						if (isNaN(c)) c = 0;
						// alert(c);

						// alert(refcontrolid);
						var controlid = data1["2"];
						// alert(controlid);




						if (data1["3"] != null) {
							//   alert(data1["3"]);
							var sum = 0;
							// alert("sum");
							if (data1["3"].indexOf("$") > -1) {
								//alert('$');
								var data2 = data1["3"].split('$')
								for (i = 1; i <= data2["0"]; i++) {
									// alert('b');
									// alert(controlid + "_" + i);
									var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}
								var cid = data2["1"];
								for (i = 1; i <= data1["4"]; i++) {
									// alert('b');
									// alert(cid + "_" + i);
									var s = document.getElementsByClassName(cid + "_" + i)[0].value;
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}

							}

							else {
								// alert('b');
								for (i = 1; i <= data1["3"]; i++) {
									// alert('b');
									// alert(controlid + "_" + i);
									var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}
							}

							// alert(sum)
							if (c != sum) {
								// alert("failcheckreference");
								var reflabelorderid = "lblorderno_" + data1["1"];
								var relabelorderid = "lblorderno_" + data1["2"];
								var x = document.getElementById(reflabelorderid).innerHTML;
								var y = document.getElementById(relabelorderid).innerHTML;
								document.getElementById("lblerror" + controlid).innerHTML = "The total of question" + y + " should equal to question " + x;

								e.preventDefault();

								// alert("a1");
								if (error.length == 0) {
									// alert(controlid+"_1");
									document.getElementsByClassName(controlid + "_1")[0].focus();
								}
								error.push("Q" + y.split('.')[0]);
								//return false;
							}

							else {
								document.getElementById("lblerror" + controlid).innerHTML = "";
							}

						}
					}


					if (object.value == "ChkTotal") {

						// alert("ChkTotal")
						var total = data1["2"];

						// alert(total);
						var controlid = data1["1"];
						if (data1["3"] != null) {
							var sum = 0;
							for (i = 1; i <= data1["3"]; i++) {
								//alert('b');
								// alert(controlid + "_" + i);
								//  var s = document.getElementsByClassName(controlid + "_" + i).value;
								var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
								// alert(s);
								s = parseInt(s);
								if (isNaN(s)) s = 0;
								sum += s;
							}

							// alert(sum);
							if (total != sum) {
								// alert("failtotal");
								document.getElementById("lblerror" + controlid).innerHTML = "Your current total of " + sum + " does not add up to the required total of 100 for question" + document.getElementById("lblorderno_" + controlid).innerHTML;

								e.preventDefault();


								// alert("a2");

								if (error.length == 0) {
									// alert(controlid+"_1");
									document.getElementsByClassName(controlid + "_1")[0].focus();
								}

								error.push("Q" + document.getElementById("lblorderno_" + controlid).innerHTML.split('.')[0]);

								//alert(error.length);
								//return false;
							}

							else {

								document.getElementById("lblerror" + controlid).innerHTML = "";

							}

						}

					}


					if (object.value == "percentagechk") {

						var controlpercentage = data1["1"];
						var spercentage = parseInt(document.getElementsByClassName(controlpercentage)[0].value);
						if (isNaN(spercentage)) {
							spercentage = 0
							//alert("value");
						}
						if (spercentage > 100) {

							document.getElementById("lblerror" + controlpercentage).innerHTML = "Max value should not be greater than  100";
							e.preventDefault();


							// alert("a3");
							if (error.length == 0) {
								// alert(controlpercentage);
								document.getElementsByClassName(controlpercentage)[0].focus();
							}

							error.push(controlpercentage.split('.')[0]);
						}

						else {
							document.getElementById("lblerror" + controlpercentage).innerHTML = "";

						}

						// alert(spercentage);
					}



					if (error.length > 0) {
						e.preventDefault();
						document.getElementById("parenterrorlbl").style.display = "block";
						document.getElementById("parenterrorlbl").innerHTML = "Please enter a valid data for " + error.join(",") + ".";
					}

					else {
						document.getElementById("parenterrorlbl").style.display = "none";
						document.getElementById("parenterrorlbl").innerHTML = "";
						return true;
					}

					//alert(error.length);
					// document.getElementById("parenterrorlbl").style.display = "block";
					// document.getElementById("parenterrorlbl").innerHTML = "Please enter a valid data for " + error.join(",")+".";

				});

			});
		});


	</script>

	<script type="text/javascript">


		$(document).ready(function () {

			document.getElementById("parenterrorlbl").innerHTML = "";

			$("#btnSubmit").click(function (e) {
				document.getElementById("parenterrorlbl").innerHTML = "";
				document.getElementById("lblsaveerror").innerHTML = "";
				document.getElementById("parenterrorlbl").style.display = "none";
				document.getElementById("lblsaveerror").style.display = "none";
				var error = new Array();

				var valemail = document.getElementsByClassName("Q37")[0].value;

				var i = 0;
				if (valemail == "") {


					document.getElementById("lblerrorQ37").innerHTML = "Please enter Email Addressbtnsave";
					error.push("Q72");
					document.getElementsByClassName("Q37")[0].focus();

					i++;

				}
				else {
					document.getElementById("lblerrorQ37").innerHTML = "";
				}

				var valname = document.getElementsByClassName("Q38")[0].value;

				if (valname == "") {

					// alert("success");
					document.getElementById("lblerrorQ38").innerHTML = "Please enter Name";

					if (error.length == 0) {

						document.getElementsByClassName("Q38")[0].focus();
					}
					error.push("Q73");
					//alert(error.length);

					//return false;

					i++;

				}
				else {

					document.getElementById("lblerrorQ38").innerHTML = "";
				}

				var valpname = document.getElementsByClassName("Q47")[0].value;

				if (valpname == "") {

					// alert("success");
					document.getElementById("lblerrorQ47").innerHTML = "Please enter Practice Name";

					if (error.length == 0) {

						document.getElementsByClassName("Q47")[0].focus();
					}

					error.push("Q74");
					//return false;

					i++;

				}
				else {

					document.getElementById("lblerrorQ47").innerHTML = "";
				}


				var street = document.getElementsByClassName("Q48_1")[0].value;
				var city = document.getElementsByClassName("Q48_2")[0].value;
				var state = document.getElementsByClassName("Q48_3")[0].value;
				var zip = document.getElementsByClassName("Q48_4")[0].value;


				if (street == "" || city == "" || state == "" || zip == "") {

					//alert("success");
					document.getElementById("lblerrorQ48").innerHTML = "Please enter Address";

					if (error.length == 0) {

						document.getElementsByClassName("Q48_1")[0].focus();
					}
					//return false;
					error.push("Q75");
					i++;

				}
				else {

					document.getElementById("lblerrorQ48").innerHTML = "";
				}



				var valphone = document.getElementsByClassName("Q49")[0].value;

				if (valphone == "") {

					// alert("success");
					document.getElementById("lblerrorQ49").innerHTML = "Please enter Phone No";

					if (error.length == 0) {

						document.getElementsByClassName("Q49")[0].focus();
					}
					error.push("Q76");
					//return false;

					i++;

				}
				else {

					document.getElementById("lblerrorQ49").innerHTML = "";
				}


				var error = new Array();
				//alert("1");
				var data = $("input:hidden[id^='hdq']");
				//alert("2");
				$.each(data, function (key, object) {

					//alert("3");
					var completestring = object.id;
					//alert(completestring.split('/').length);
					var data1 = completestring.split('/')
					if (object.value == "ChkQuesRef1" || object.value == "ChQuesRef2") {

						// alert(object.value)
						var refcontrolid = data1["1"];
						var c = document.getElementsByClassName(refcontrolid)[0].value;
						c = parseInt(c);
						if (isNaN(c)) c = 0;
						// alert(c);

						// alert(refcontrolid);
						var controlid = data1["2"];
						// alert(controlid);




						if (data1["3"] != null) {
							//   alert(data1["3"]);
							var sum = 0;
							// alert("sum");
							if (data1["3"].indexOf("$") > -1) {
								//alert('$');
								var data2 = data1["3"].split('$')
								for (j = 1; j <= data2["0"]; j++) {
									// alert('b');
									// alert(controlid + "_" + i);
									var s = document.getElementsByClassName(controlid + "_" + j)[0].value;
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}
								var cid = data2["1"];
								for (j = 1; j <= data1["4"]; j++) {
									// alert('b');
									// alert(cid + "_" + i);
									var s = document.getElementsByClassName(cid + "_" + j)[0].value;
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}

							}

							else {
								// alert('b');
								for (j = 1; j <= data1["3"]; j++) {
									// alert('b');
									// alert(controlid + "_" + i);
									var s = document.getElementsByClassName(controlid + "_" + j)[0].value;
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}
							}

							// alert(sum)
							if (c != sum) {
								// alert("failcheckreference");
								var reflabelorderid = "lblorderno_" + data1["1"];
								var relabelorderid = "lblorderno_" + data1["2"];
								var x = document.getElementById(reflabelorderid).innerHTML;
								var y = document.getElementById(relabelorderid).innerHTML;
								document.getElementById("lblerror" + controlid).innerHTML = "The total of question" + y + " should equal to question " + x;

								e.preventDefault();

								// alert("a1");
								if (error.length == 0) {
									// alert(controlid+"_1");
									document.getElementsByClassName(controlid + "_1")[0].focus();
								}
								error.push("Q" + y.split('.')[0]);
								//return false;
							}

							else {
								document.getElementById("lblerror" + controlid).innerHTML = "";
							}

						}
					}


					if (object.value == "ChkTotal") {

						// alert("ChkTotal")
						var total = data1["2"];

						// alert(total);
						var controlid = data1["1"];
						if (data1["3"] != null) {
							var sum = 0;
							for (j = 1; j <= data1["3"]; j++) {
								//alert('b');
								// alert(controlid + "_" + i);
								//  var s = document.getElementsByClassName(controlid + "_" + i).value;
								var s = document.getElementsByClassName(controlid + "_" + j)[0].value;
								// alert(s);
								s = parseInt(s);
								if (isNaN(s)) s = 0;
								sum += s;
							}

							// alert(sum);
							if (total != sum) {
								//alert("failtotal");
								document.getElementById("lblerror" + controlid).innerHTML = "Your current total of " + sum + " does not add up to the required total of 100 for question" + document.getElementById("lblorderno_" + controlid).innerHTML;

								e.preventDefault();


								// alert("a2");

								if (error.length == 0) {
									//alert(controlid + "_1");
									document.getElementsByClassName(controlid + "_1")[0].focus();
								}

								error.push("Q" + document.getElementById("lblorderno_" + controlid).innerHTML.split('.')[0]);

								//alert(error.length);
								//return false;
							}

							else {

								document.getElementById("lblerror" + controlid).innerHTML = "";

							}

						}

					}


					if (object.value == "percentagechk") {

						var controlpercentage = data1["1"];
						var spercentage = parseInt(document.getElementsByClassName(controlpercentage)[0].value);
						if (isNaN(spercentage)) {
							spercentage = 0
							//alert("value");
						}
						if (spercentage > 100) {

							document.getElementById("lblerror" + controlpercentage).innerHTML = "Max value should not be greater than  100";
							e.preventDefault();


							// alert("a3");
							if (error.length == 0) {
								// alert(controlpercentage);
								document.getElementsByClassName(controlpercentage)[0].focus();
							}

							error.push(controlpercentage.split('.')[0]);
						}

						else {
							document.getElementById("lblerror" + controlpercentage).innerHTML = "";

						}

						// alert(spercentage);
					}
					document.getElementById("parenterrorlbl").style.display = "block";
					document.getElementById("parenterrorlbl").innerHTML = "Survey could not be saved due to error in " + error.join(",") + ". Please enter valid data.";

					//if (error.length > 0) {
					//                           document.getElementById("lblsaveerror").style.display = "none";

					//}

					//else
					//{
					//    document.getElementById("parenterrorlbl").style.display = "none";
					//    document.getElementById("lblsaveerror").style.display = "block";
					//}

				});



				if (i > 0) {
					e.preventDefault();
					document.getElementById("parenterrorlbl").style.display = "block";
					document.getElementById("parenterrorlbl").innerHTML = "Error occured for " + error.join(",");
				}

				else {
					document.getElementById("parenterrorlbl").style.display = "none";
					document.getElementById("parenterrorlbl").innerHTML = "";
					return true;
				}



			});
		});

	</script>

	<script type="text/javascript">


		$(document).ready(function () {
			//document.getElementById("lblsaveerror").innerHTML = "";
			document.getElementById("parenterrorlbl").innerHTML = "";
			$("#btnNextbottom").click(function (e) {
				document.getElementById("parenterrorlbl").innerHTML = "";

				document.getElementById("lblsaveerror").innerHTML = "";
				document.getElementById("parenterrorlbl").style.display = "none";
				document.getElementById("lblsaveerror").style.display = "none";
				//ALERT("TEST");
				//var cpercen = document.getElementsByClassName("Q41")[0].value;
				//ALERT(cpercen);
				//var cpercen = document.getElementsByClassName("Q41")[0].value;
				//  alert(document.getElementsByClassName("Q41")[0].value);
				//alert('t1');
				//alert('t');

				var error = new Array();

				var data = $("input:hidden[id^='hdq']");
				$.each(data, function (key, object) {
					// alert(object.value);
					//alert(object.id);
					var completestring = object.id;
					//alert(completestring.split('/').length);
					var data1 = completestring.split('/')
					if (object.value == "ChkQuesRef1" || object.value == "ChQuesRef2") {

						// alert(object.value)
						var refcontrolid = data1["1"];
						var c = document.getElementsByClassName(refcontrolid)[0].value;
						c = parseInt(c);
						if (isNaN(c)) c = 0;
						// alert(c);

						// alert(refcontrolid);
						var controlid = data1["2"];
						// alert(controlid);




						if (data1["3"] != null) {
							//   alert(data1["3"]);
							var sum = 0;
							// alert("sum");
							if (data1["3"].indexOf("$") > -1) {
								//alert('$');
								var data2 = data1["3"].split('$')
								for (i = 1; i <= data2["0"]; i++) {
									// alert('b');
									// alert(controlid + "_" + i);
									var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}
								var cid = data2["1"];
								for (i = 1; i <= data1["4"]; i++) {
									// alert('b');
									// alert(cid + "_" + i);
									var s = document.getElementsByClassName(cid + "_" + i)[0].value;
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}

							}

							else {
								// alert('b');
								for (i = 1; i <= data1["3"]; i++) {
									// alert('b');
									// alert(controlid + "_" + i);
									var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}
							}

							// alert(sum)
							if (c != sum) {
								// alert("failcheckreference");
								var reflabelorderid = "lblorderno_" + data1["1"];
								var relabelorderid = "lblorderno_" + data1["2"];
								var x = document.getElementById(reflabelorderid).innerHTML;
								var y = document.getElementById(relabelorderid).innerHTML;
								document.getElementById("lblerror" + controlid).innerHTML = "The total of question" + y + " should equal to question " + x;

								e.preventDefault();

								//alert("a1");
								if (error.length == 0) {
									//alert(controlid + "_1");
									document.getElementsByClassName(controlid + "_1")[0].focus();
								}
								error.push("Q" + y.split('.')[0]);
								//return false;
							}

							else {
								document.getElementById("lblerror" + controlid).innerHTML = "";
							}

						}
					}


					if (object.value == "ChkTotal") {

						// alert("ChkTotal")
						var total = data1["2"];

						// alert(total);
						var controlid = data1["1"];
						if (data1["3"] != null) {
							var sum = 0;
							for (i = 1; i <= data1["3"]; i++) {
								//alert('b');
								// alert(controlid + "_" + i);
								//  var s = document.getElementsByClassName(controlid + "_" + i).value;
								var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
								// alert(s);
								s = parseInt(s);
								if (isNaN(s)) s = 0;
								sum += s;
							}

							// alert(sum);
							if (total != sum) {
								// alert("failtotal");
								document.getElementById("lblerror" + controlid).innerHTML = "Your current total of " + sum + " does not add up to the required total of 100 for question" + document.getElementById("lblorderno_" + controlid).innerHTML;

								e.preventDefault();


								// alert("a2");

								if (error.length == 0) {
									//alert(controlid + "_1");
									document.getElementsByClassName(controlid + "_1")[0].focus();
								}

								error.push("Q" + document.getElementById("lblorderno_" + controlid).innerHTML.split('.')[0]);

								//alert(error.length);
								//return false;
							}

							else {

								document.getElementById("lblerror" + controlid).innerHTML = "";

							}

						}

					}


					if (object.value == "percentagechk") {

						var controlpercentage = data1["1"];
						var spercentage = parseInt(document.getElementsByClassName(controlpercentage)[0].value);
						if (isNaN(spercentage)) {
							spercentage = 0
							//alert("value");
						}
						if (spercentage > 100) {

							document.getElementById("lblerror" + controlpercentage).innerHTML = "Max value should not be greater than  100";
							e.preventDefault();


							//alert("a3");
							if (error.length == 0) {
								// alert(controlpercentage);
								document.getElementsByClassName(controlpercentage)[0].focus();
							}

							error.push(controlpercentage.split('.')[0]);
						}

						else {
							document.getElementById("lblerror" + controlpercentage).innerHTML = "";

						}

						// alert(spercentage);
					}

					//alert(error.length);
					document.getElementById("parenterrorlbl").style.display = "block";
					document.getElementById("parenterrorlbl").innerHTML = "Please enter a valid data for " + error.join(",") + ".";

				});

			});
		});


	</script>

	<script type="text/javascript">


			$(document).ready(function () {

				document.getElementById("parenterrorlbl").innerHTML = "";
				$("#btnSave").click(function (e) {
					document.getElementById("parenterrorlbl").innerHTML = "";

					document.getElementById("parenterrorlbl").style.display = "none";
					document.getElementById("lblsaveerror").innerHTML = "";
					document.getElementById("lblsaveerror").style.display = "none";
					//alert("save");

					var error = new Array();
					//alert("1");
					var data = $("input:hidden[id^='hdq']");
					//alert("2");
					$.each(data, function (key, object) {

						//alert("3");
						var completestring = object.id;
						//alert(completestring.split('/').length);
						var data1 = completestring.split('/')
						if (object.value == "ChkQuesRef1" || object.value == "ChQuesRef2") {

							// alert(object.value)
							var refcontrolid = data1["1"];
							var c = document.getElementsByClassName(refcontrolid)[0].value;
							c = parseInt(c);
							if (isNaN(c)) c = 0;
							// alert(c);

							// alert(refcontrolid);
							var controlid = data1["2"];
							// alert(controlid);




							if (data1["3"] != null) {
								//   alert(data1["3"]);
								var sum = 0;
								// alert("sum");
								if (data1["3"].indexOf("$") > -1) {
									//alert('$');
									var data2 = data1["3"].split('$')
									for (i = 1; i <= data2["0"]; i++) {
										// alert('b');
										// alert(controlid + "_" + i);
										var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
										s = parseInt(s);
										if (isNaN(s)) s = 0;
										sum += s;
									}
									var cid = data2["1"];
									for (i = 1; i <= data1["4"]; i++) {
										// alert('b');
										// alert(cid + "_" + i);
										var s = document.getElementsByClassName(cid + "_" + i)[0].value;
										s = parseInt(s);
										if (isNaN(s)) s = 0;
										sum += s;
									}

								}

								else {
									// alert('b');
									for (i = 1; i <= data1["3"]; i++) {
										// alert('b');
										// alert(controlid + "_" + i);
										var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
										s = parseInt(s);
										if (isNaN(s)) s = 0;
										sum += s;
									}
								}

								// alert(sum)
								if (c != sum) {
									// alert("failcheckreference");
									var reflabelorderid = "lblorderno_" + data1["1"];
									var relabelorderid = "lblorderno_" + data1["2"];
									var x = document.getElementById(reflabelorderid).innerHTML;
									var y = document.getElementById(relabelorderid).innerHTML;
									document.getElementById("lblerror" + controlid).innerHTML = "The total of question" + y + " should equal to question " + x;

									e.preventDefault();

									// alert("a1");
									if (error.length == 0) {
										// alert(controlid+"_1");
										document.getElementsByClassName(controlid + "_1")[0].focus();
									}
									error.push("Q" + y.split('.')[0]);
									//return false;
								}

								else {
									document.getElementById("lblerror" + controlid).innerHTML = "";
								}

							}
						}


						if (object.value == "ChkTotal") {

							// alert("ChkTotal")
							var total = data1["2"];

							// alert(total);
							var controlid = data1["1"];
							if (data1["3"] != null) {
								var sum = 0;
								for (i = 1; i <= data1["3"]; i++) {
									//alert('b');
									// alert(controlid + "_" + i);
									//  var s = document.getElementsByClassName(controlid + "_" + i).value;
									var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
									// alert(s);
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}

								// alert(sum);
								if (total != sum) {
									//alert("failtotal");
									document.getElementById("lblerror" + controlid).innerHTML = "Your current total of " + sum + " does not add up to the required total of 100 for question" + document.getElementById("lblorderno_" + controlid).innerHTML;

									e.preventDefault();


									// alert("a2");

									if (error.length == 0) {
										//alert(controlid + "_1");
										document.getElementsByClassName(controlid + "_1")[0].focus();
									}

									error.push("Q" + document.getElementById("lblorderno_" + controlid).innerHTML.split('.')[0]);

									//alert(error.length);
									//return false;
								}

								else {

									document.getElementById("lblerror" + controlid).innerHTML = "";

								}

							}

						}


						if (object.value == "percentagechk") {

							var controlpercentage = data1["1"];
							var spercentage = parseInt(document.getElementsByClassName(controlpercentage)[0].value);
							if (isNaN(spercentage)) {
								spercentage = 0
								//alert("value");
							}
							if (spercentage > 100) {

								document.getElementById("lblerror" + controlpercentage).innerHTML = "Max value should not be greater than  100";
								e.preventDefault();


								// alert("a3");
								if (error.length == 0) {
									// alert(controlpercentage);
									document.getElementsByClassName(controlpercentage)[0].focus();
								}

								error.push(controlpercentage.split('.')[0]);
							}

							else {
								document.getElementById("lblerror" + controlpercentage).innerHTML = "";

							}

							// alert(spercentage);
						}
						document.getElementById("parenterrorlbl").style.display = "block";
						document.getElementById("parenterrorlbl").innerHTML = "Survey could not be saved due to error in " + error.join(",") + ". Please enter valid data.";

						//if (error.length > 0) {
						//                           document.getElementById("lblsaveerror").style.display = "none";

						//}

						//else
						//{
						//    document.getElementById("parenterrorlbl").style.display = "none";
						//    document.getElementById("lblsaveerror").style.display = "block";
						//}

					});


					var lasterror = new Array();
					var valemail = document.getElementsByClassName("Q37")[0];
					//alert('test1');
					var i1 = 0;
					if (valemail === undefined || valemail === null) {

						//alert(valemail);

					}

					else {

						varemailvalue = document.getElementsByClassName("Q37")[0].value;
						//alert(varemailvalue);

						//alert(varemailvalue);
						if (varemailvalue == "") {
							//alert('in');
							document.getElementById("lblerrorQ37").innerHTML = "Please enter Email Address";
							//alert(document.getElementById("lblerrorQ37"));
							lasterror.push("Q72");
							document.getElementsByClassName("Q37")[0].focus();

							i1++;

						}
						else {
							document.getElementById("lblerrorQ37").innerHTML = "";
						}

					}


					var valname = document.getElementsByClassName("Q38")[0];


					if (valname === undefined || valname === null) {

					}

					else {

						valnamevalue = document.getElementsByClassName("Q38")[0].value;

						//alert(valnamevalue);
						if (valnamevalue == "") {
							//alert('in');
							document.getElementById("lblerrorQ38").innerHTML = "Please enter Name";
							//alert(document.getElementById("lblerrorQ38"));



							if (lasterror.length == 0) {

								document.getElementsByClassName("Q38")[0].focus();
							}
							lasterror.push("Q73");

							i1++;

						}
						else {
							document.getElementById("lblerrorQ38").innerHTML = "";
						}

					}


					var valpname = document.getElementsByClassName("Q47")[0];


					if (valpname === undefined || valpname === null) {

					}

					else {

						valpnamevalue = document.getElementsByClassName("Q47")[0].value;


						if (valpnamevalue == "") {
							//alert('in');
							document.getElementById("lblerrorQ47").innerHTML = "Please enter Practice  Name";




							if (lasterror.length == 0) {

								document.getElementsByClassName("Q47")[0].focus();
							}
							lasterror.push("Q47");

							i1++;

						}
						else {
							document.getElementById("lblerrorQ47").innerHTML = "";
						}

					}


					var street = document.getElementsByClassName("Q48_1")[0];
					var city = document.getElementsByClassName("Q48_2")[0];
					var state = document.getElementsByClassName("Q48_3")[0];
					var zip = document.getElementsByClassName("Q48_4")[0];
					if ((street === undefined || street === null) && (city === undefined || city === null) && (state === undefined || state === null) && (zip === undefined || zip === null)) {

					}

					else {


						var streetvalue = document.getElementsByClassName("Q48_1")[0].value;
						var cityvalue = document.getElementsByClassName("Q48_2")[0].value;
						var statevalue = document.getElementsByClassName("Q48_3")[0].value;
						var zipvalue = document.getElementsByClassName("Q48_4")[0].value;

						if (streetvalue == "" || cityvalue == "" || statevalue == "" || zipvalue == "") {

							//alert("success");
							document.getElementById("lblerrorQ48").innerHTML = "Please enter Address";

							if (lasterror.length == 0) {

								document.getElementsByClassName("Q48_1")[0].focus();
							}
							//return false;
							lasterror.push("Q75");
							i1++;

						}
						else {

							document.getElementById("lblerrorQ48").innerHTML = "";
						}

					}



					var valphone = document.getElementsByClassName("Q49")[0];


					if (valphone === undefined || valphone === null) {

					}

					else {

						valphonevalue = document.getElementsByClassName("Q49")[0].value;


						if (valphonevalue == "") {
							//alert('in');
							document.getElementById("lblerrorQ49").innerHTML = "Please enter Phone";




							if (lasterror.length == 0) {

								document.getElementsByClassName("Q49")[0].focus();
							}
							lasterror.push("Q76");

							i1++;

						}
						else {
							document.getElementById("lblerrorQ49").innerHTML = "";
						}

					}


					if (i1 > 0) {
						e.preventDefault();
						document.getElementById("parenterrorlbl").style.display = "block";
						document.getElementById("parenterrorlbl").innerHTML = "Error occured for " + lasterror.join(",");

					}

					else {

						return true;
					}




				});
			});


	</script>

	<script type="text/javascript">
			function callMethod() {
				// alert('test');
				document.getElementById("lblfirst").innerHTML = "";
				var name = document.getElementById("txtname").value;
				document.getElementById("lblfirst").style.display = "none";
				var practice = document.getElementById("txtpractice").value;

				//alert(name);

				// alert(practice);
				var error = new Array();
				var i = 0;
				if (name.trim() == "") {

					//alert("enter a name");
					error.push("Name");
					i++;
				}

				if (practice.trim() == "") {

					//alert("enter a practice");
					error.push("Practice Name");
					i++
				}

				if (i > 0) {
					document.getElementById("lblfirst").innerHTML = "Please Enter value for " + error.join(",");
					document.getElementById("lblfirst").style.display = "block";
					return false;
				}

				else {
					document.getElementById("lblfirst").innerHTML = "";
					return true;
				}

				//return false;
			}
	</script>

	<script type="text/javascript">


			$(document).ready(function () {

				document.getElementById("parenterrorlbl").innerHTML = "";

				$("#btnPreview").click(function (e) {
					document.getElementById("parenterrorlbl").innerHTML = "";
					document.getElementById("lblsaveerror").innerHTML = "";
					document.getElementById("parenterrorlbl").style.display = "none";
					document.getElementById("lblsaveerror").style.display = "none";
					var error = new Array();
					var data = $("input:hidden[id^='hdq']");
					$.each(data, function (key, object) {

						//alert("3");
						var completestring = object.id;
						//alert(completestring.split('/').length);
						var data1 = completestring.split('/')
						if (object.value == "ChkQuesRef1" || object.value == "ChQuesRef2") {

							// alert(object.value)
							var refcontrolid = data1["1"];
							var c = document.getElementsByClassName(refcontrolid)[0].value;
							c = parseInt(c);
							if (isNaN(c)) c = 0;
							// alert(c);

							// alert(refcontrolid);
							var controlid = data1["2"];
							// alert(controlid);




							if (data1["3"] != null) {
								//   alert(data1["3"]);
								var sum = 0;
								// alert("sum");
								if (data1["3"].indexOf("$") > -1) {
									//alert('$');
									var data2 = data1["3"].split('$')
									for (i = 1; i <= data2["0"]; i++) {
										// alert('b');
										// alert(controlid + "_" + i);
										var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
										s = parseInt(s);
										if (isNaN(s)) s = 0;
										sum += s;
									}
									var cid = data2["1"];
									for (i = 1; i <= data1["4"]; i++) {
										// alert('b');
										// alert(cid + "_" + i);
										var s = document.getElementsByClassName(cid + "_" + i)[0].value;
										s = parseInt(s);
										if (isNaN(s)) s = 0;
										sum += s;
									}

								}

								else {
									// alert('b');
									for (i = 1; i <= data1["3"]; i++) {
										// alert('b');
										// alert(controlid + "_" + i);
										var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
										s = parseInt(s);
										if (isNaN(s)) s = 0;
										sum += s;
									}
								}

								// alert(sum)
								if (c != sum) {
									// alert("failcheckreference");
									var reflabelorderid = "lblorderno_" + data1["1"];
									var relabelorderid = "lblorderno_" + data1["2"];
									var x = document.getElementById(reflabelorderid).innerHTML;
									var y = document.getElementById(relabelorderid).innerHTML;
									document.getElementById("lblerror" + controlid).innerHTML = "The total of question" + y + " should equal to question " + x;

									e.preventDefault();

									// alert("a1");
									if (error.length == 0) {
										// alert(controlid+"_1");
										document.getElementsByClassName(controlid + "_1")[0].focus();
									}
									error.push("Q" + y.split('.')[0]);
									//return false;
								}

								else {
									document.getElementById("lblerror" + controlid).innerHTML = "";
								}

							}
						}


						if (object.value == "ChkTotal") {

							// alert("ChkTotal")
							var total = data1["2"];

							// alert(total);
							var controlid = data1["1"];
							if (data1["3"] != null) {
								var sum = 0;
								for (i = 1; i <= data1["3"]; i++) {
									//alert('b');
									// alert(controlid + "_" + i);
									//  var s = document.getElementsByClassName(controlid + "_" + i).value;
									var s = document.getElementsByClassName(controlid + "_" + i)[0].value;
									// alert(s);
									s = parseInt(s);
									if (isNaN(s)) s = 0;
									sum += s;
								}

								// alert(sum);
								if (total != sum) {
									//alert("failtotal");
									document.getElementById("lblerror" + controlid).innerHTML = "Your current total of " + sum + " does not add up to the required total of 100 for question" + document.getElementById("lblorderno_" + controlid).innerHTML;

									e.preventDefault();


									// alert("a2");

									if (error.length == 0) {
										//alert(controlid + "_1");
										document.getElementsByClassName(controlid + "_1")[0].focus();
									}

									error.push("Q" + document.getElementById("lblorderno_" + controlid).innerHTML.split('.')[0]);

									//alert(error.length);
									//return false;
								}

								else {

									document.getElementById("lblerror" + controlid).innerHTML = "";

								}

							}

						}


						if (object.value == "percentagechk") {

							var controlpercentage = data1["1"];
							var spercentage = parseInt(document.getElementsByClassName(controlpercentage)[0].value);
							if (isNaN(spercentage)) {
								spercentage = 0
								//alert("value");
							}
							if (spercentage > 100) {

								document.getElementById("lblerror" + controlpercentage).innerHTML = "Max value should not be greater than  100";
								e.preventDefault();


								// alert("a3");
								if (error.length == 0) {
									// alert(controlpercentage);
									document.getElementsByClassName(controlpercentage)[0].focus();
								}

								error.push(controlpercentage.split('.')[0]);
							}

							else {
								document.getElementById("lblerror" + controlpercentage).innerHTML = "";

							}

							// alert(spercentage);
						}
						document.getElementById("parenterrorlbl").style.display = "block";
						document.getElementById("parenterrorlbl").innerHTML = "Survey could not be saved due to error in " + error.join(",") + ". Please enter valid data.";

						//if (error.length > 0) {
						//                           document.getElementById("lblsaveerror").style.display = "none";

						//}

						//else
						//{
						//    document.getElementById("parenterrorlbl").style.display = "none";
						//    document.getElementById("lblsaveerror").style.display = "block";
						//}

					});
					var valemail = document.getElementsByClassName("Q37")[0].value;

					var i = 0;
					if (valemail == "") {


						document.getElementById("lblerrorQ37").innerHTML = "Please enter Email Address";
						error.push("Q72");
						document.getElementsByClassName("Q37")[0].focus();

						i++;

					}
					else {
						document.getElementById("lblerrorQ37").innerHTML = "";
					}

					var valname = document.getElementsByClassName("Q38")[0].value;

					if (valname == "") {

						// alert("success");
						document.getElementById("lblerrorQ38").innerHTML = "Please enter Name";

						if (error.length == 0) {

							document.getElementsByClassName("Q38")[0].focus();
						}
						error.push("Q73");
						//alert(error.length);

						//return false;

						i++;

					}
					else {

						document.getElementById("lblerrorQ38").innerHTML = "";
					}

					var valpname = document.getElementsByClassName("Q47")[0].value;

					if (valpname == "") {

						// alert("success");
						document.getElementById("lblerrorQ47").innerHTML = "Please enter Practice Name";

						if (error.length == 0) {

							document.getElementsByClassName("Q47")[0].focus();
						}

						error.push("Q74");
						//return false;

						i++;

					}
					else {

						document.getElementById("lblerrorQ47").innerHTML = "";
					}


					var street = document.getElementsByClassName("Q48_1")[0].value;
					var city = document.getElementsByClassName("Q48_2")[0].value;
					var state = document.getElementsByClassName("Q48_3")[0].value;
					var zip = document.getElementsByClassName("Q48_4")[0].value;


					if (street == "" || city == "" || state == "" || zip == "") {

						//alert("success");
						document.getElementById("lblerrorQ48").innerHTML = "Please enter Address";

						if (error.length == 0) {

							document.getElementsByClassName("Q48_1")[0].focus();
						}
						//return false;
						error.push("Q75");
						i++;

					}
					else {

						document.getElementById("lblerrorQ48").innerHTML = "";
					}



					var valphone = document.getElementsByClassName("Q49")[0].value;

					if (valphone == "") {

						// alert("success");
						document.getElementById("lblerrorQ49").innerHTML = "Please enter Phone No";

						if (error.length == 0) {

							document.getElementsByClassName("Q49")[0].focus();
						}
						error.push("Q76");
						//return false;

						i++;

					}
					else {

						document.getElementById("lblerrorQ49").innerHTML = "";
					}

					if (i > 0) {
						e.preventDefault();
						document.getElementById("parenterrorlbl").style.display = "block";
						document.getElementById("parenterrorlbl").innerHTML = "Error occured for " + error.join(",");
					}

					else {

						return true;
					}
					//    if (pgnd == "" || pgng == 'undefined' || pgnd== null) {
					//        alert('The textbox should not be empty...');

					//        document.getElementsByClassName("lblerrorQ49").value = "Please enter a Email";
					//        //document.getElementById("Q49").focus();
					//    return false;
					//}


				});
			});

	</script>
	<title></title>
</head>
<body>

	<form runat="server">

		<div class="ThemeWrapper">
			<div id="nav_princ">
				<div id="nav">
					<div id="logo">
						<a id="ctl00_PlaceHolderTopNavBar_LogoHomePageLink" title="Home page" href="http://www.essilor.com/en">
							<img title="Home page" src="images/head_logoV2.png" alt=""></a>
					</div>
					<div id="nav_elements">
						<div id="menu_constant">

							<div id="insitu">
								<a id="ctl00_PlaceHolderTopNavBar_ctl00_HomePageLink" href="http://www.essilor.com/en" style="display: none">Home page</a>


								<script type="text/javascript">
									function hover(element, imgUrl) {
										imgUrl = imgUrl.split('.')[0] + 'hover.' + imgUrl.split('.')[1];
										element.setAttribute('src', imgUrl);
									}
									function unhover(element, imgUrl) {
										element.setAttribute('src', imgUrl);
									}
								</script>
								<div id="SocialMediaLinks" style="overflow: auto; float: left;">

									<a href="http://www.facebook.com/pages/Essilor/88196063119" target="_blank">
										<img style="height: 14px; width: 14px;" src="images/facebookV2.png" alt="facebook" onmouseover="hover(this,'images/facebookV2.png');" onmouseout="unhover(this,'images/facebookV2.png');"></a>

									<a href="https://twitter.com/#!/Essilor" target="_blank">
										<img style="height: 14px; width: 14px;" src="images/twitterV2.png" alt="Twitter" onmouseover="hover(this,'images/twitterV2.png');" onmouseout="unhover(this,'images/twitterV2.png');"></a>

									<a href="http://www.youtube.com/EssilorCorp" target="_blank">
										<img style="height: 14px; width: 14px;" src="images/youtubeV2.png" alt="Youtube" onmouseover="hover(this,'images/youtubeV2.png');" onmouseout="unhover(this,'images/youtubeV2.png');"></a>

									<a href="http://www.linkedin.com/company/essilor" target="_blank">
										<img style="height: 14px; width: 14px;" src="images/linkedinV2.png" alt="linkedind" onmouseover="hover(this,'images/linkedinV2.png');" onmouseout="unhover(this,'images/linkedinV2.png');"></a>

								</div>


							</div>
							<div id="menu_headV2">
								<ul>
									<li>Language : 
       <a class="nav_lang exe" href="http://www.essilor.com/en/Pages/Home.aspx">en</a>

										<a class="nav_lang" href="http://www.essilor.com/fr">fr</a>
									</li>
									<li><a href="/en/Pages/Registration.aspx">Email alerts</a></li>
									<li><a href="/en/Pages/Registration.aspx?feeds=1">RSS feeds</a></li>
									<li><a href="/en/Press/Pages/Home.aspx">Media Library</a></li>
								</ul>
								<input type="text" id="search" class="SearchInput" name="SearchString" accesskey="S" value="Search..." onfocus="if (this.value=='Search...'){this.value=''}" onblur="if (this.value==''){this.value ='Search...';}" title="Search"><a id="ok_search" class="OkSearchBtn" href="javascript:" onclick="javascript:CustomSearchRedirect('/en/pages/search.aspx');javascript:return false;">Ok</a>
							</div>


						</div>
						<div id="menu_nav">

							<!-- Top Navigation Menu -->
							<div id="nav_blackV2">

								<a href="/en/Group/Pages/Home.aspx" target="">Group</a>
								<a href="/en/Innovation/Pages/Home.aspx" target="">Innovation</a>
								<a href="/en/EyeHealth/Pages/Home.aspx" target="">Eye care</a>
								<a href="/en/BrandsAndProducts/Pages/Brandsproducts.aspx" target="">Brands &amp; products</a>




							</div>
							<div id="menu_nav_sep">
							</div>
							<div id="nav_greyV2">
								<a href="/en/Talents/Pages/Home.aspx" target="">Talents</a><a href="/en/Investors/Pages/Home.aspx" target=""> Investors</a><a style="" href="/en/Press/Pages/Default.aspx" target=""> Press</a>
							</div>
							<!-- End Top Navigation Menu -->

						</div>
					</div>
				</div>

			</div>
			<div class="bgimg">
				<table cellspacing="0" cellpadding="0" border="0" class="v-Table" style="width: 66% !important;">
					<tbody>
						<tr>
							<td class="v-TL">
								<div class="LeftPadding"></div>
							</td>
							<td class="v-TM">
								<div class="TopPadding"></div>
							</td>
							<td class="v-TR">
								<div class="LeftPadding"></div>
							</td>
						</tr>
						<tr>
							<td class="v-ML"></td>
							<td class="v-MM">
								<div id="top"><i></i></div>
								<div id="surveyBlock" style="position: relative !important;">

									<div id="surveyBlockNest">

										<div runat="server" id="firstpage" clientidmode="Static" class="mydiv">

											<asp:Label ClientIDMode="Static" ID="lblfirst" class="alert-dangerfirst" runat="server"></asp:Label>

											<span class="text-block">
												<img src="https://deloittesurvey.deloitte.com/Community/surveys/1069620006/16fca35fimg061.png" style="float: right;" alt="[Image]"><br>
											</span>




											<span class="text-block"><strong style="color: #336699"><font style="font-size: 20px;"><font color="#1f497d" style="font-size: 26px;"><u>Practice&nbsp;</u></font></font></strong><strong style="color: #336699"><font style="font-size: 20px;"><font color="#1f497d" style="font-size: 26px;"><u>Performance&nbsp;Assessment&nbsp;</u></font></font></strong><strong><font style="font-size: 20px;"><br>
												<br>
												<font style="color:#000000;font-size:18px;font-style:italic">MBA Practice Essentials - Prepared for&nbsp;Essilor <font style="font-style:italic; font-size:18px">Eye Care</font> Experts&nbsp;&nbsp;<br>
													&nbsp;</font>


												<div align="right" class="practice-label" style="padding: 10px 0">
													<div style="float: left; padding-right: 10px">
														<asp:Label runat="server" ID="lblyear">Select Survey Year:</asp:Label>
														<asp:DropDownList ID="ddlyear" runat="server"></asp:DropDownList>
													</div>
													<div style="float: left; padding-right: 10px">
														<label id="lblname" runat="server">Name:</label><asp:TextBox runat="server" ID="txtname" Style="width: 115px; border-radius: 3px; border: 1px solid #ccc;"></asp:TextBox>
													</div>
													<div style="float: left">
														<label id="lblpracticename" runat="server">Practice Name:</label><asp:TextBox runat="server" ID="txtpractice" Style="width: 115px; border-radius: 3px; border: 1px solid #ccc;"></asp:TextBox>
													</div>
													<div style="clear: both"></div>



												</div>
												<font color="#000000"><br>
														Tips</font> <br />
												<%--<div align="right" style="display: inline-block; float: right">
   <asp:Label runat="server" ID="lblyear" style="padding-left: 344px;">Select Survey Year</asp:Label>
    <asp:DropDownList  ID="ddlyear" runat="server">
                                       </asp:DropDownList>
       </div>--%>


												


											</font></strong>
												<br>
												<strong style="color: rgb(0, 0, 0); font-style:italic">Note: For best viewing results, maximize this window so it takes up the full screen.<br>
												</strong>
												<font color="#000000">&nbsp;</font><br>
												<strong style="color: rgb(0, 0, 0);"><font style="font-size: 14px;">Printing:</font></strong><br>
												<font color="#000000">At any time you may print out the survey by clicking the <strong style="color:black !important">Print</strong> button at the bottom of each page. A new window will open with the entire survey on one screen for you to print. You will still need to key in your answers to submit them, but if you need to refer to some other documents to answer some questions, you may find it helpful to print out the questionnaire. </font>
												<br>
												<br>
												<font style="color: rgb(0, 0, 0); font-size: 14px;"><strong>Entering Numerical Data:</strong><br>
												</font><font color="#000000">Some questions require that the sum of your responses to parts of the question add to 100%. The data collection software will alert you if this requirement has not been met. When reporting percentages in the questionnaire, enter them as integers, not as decimals and do not enter a % sign. For example: enter 11, not 0.11 or .11 or 11%</font><br>
												<font color="#000000">When answering numerical questions </font><strong style="color:black !important">please do not use commas</strong><font color="#000000">. For example: enter 1000 instead of 1,000.</font><br>
												<font color="#000000">When currency values are requested, report values to the nearest dollar. Do not round to the nearest thousand dollars for questions about practice revenue and expenses. Do not type in dollar signs.</font><br>
												<font color="#000000">&nbsp;</font><br>
												<strong style="color: rgb(0, 0, 0);"><font style="font-size: 14px;">Stopping and Saving mid-survey:</font></strong><br>
												<font color="#000000">At any time you may stop taking the survey and come back to it at a later time. In order to do this you must click the <b>Save </b>button at the bottom of any page. After clicking <b>Save</b>, you will be able to return to the survey at any time by re-entering your name and practice details on the survey home page.</font><br>
												<font color="#000000">&nbsp;</font><br>
												<strong style="color: rgb(0, 0, 0);"><font style="font-size: 14px;">Complete and Submit:</font></strong><br>
												<font color="#000000">Be sure to click the <b>Submit </b>button when you have completed the questionnaire.
                                                    <br>
													Once you click Submit, you cannot make any changes to the questionnaire.</font>

											</span>

											</br>
  <asp:Button ClientIDMode="Static" runat="server" ID="firstpagebutton" OnClientClick="return callMethod();" Text="Next" class="submit-button next" OnClick="firstpagebutton_Click" />
											<asp:Button runat="server" ID="firstpagebuttonbottom" ClientIDMode="Static" OnClientClick="return callMethod();" Text="Next" class="submit-button next bottom" OnClick="firstpagebutton_Click" />

											<%--<input type="submit" onclick="document.PdcSurvey.PdcButtonPressed.value='next';" value="Next" name="next" class="submit-button">  --%>
										</div>


										<div id="secondpage" runat="server" clientidmode="Static">
											<div class="clsheader">
												<asp:Label ClientIDMode="Static" ID="lblyear1" runat="server"></asp:Label>
												<asp:Label ClientIDMode="Static" ID="lblyeartext" runat="server"></asp:Label>
												<asp:Label ClientIDMode="Static" ID="lblname1" runat="server"></asp:Label>
												<asp:Label ClientIDMode="Static" ID="lblnametext" runat="server"></asp:Label>

												<asp:Label ClientIDMode="Static" ID="lblpracticename1" runat="server"></asp:Label>
												<asp:Label ClientIDMode="Static" ID="lblpracticetext" runat="server"></asp:Label>
											</div>

											<asp:Label ClientIDMode="Static" ID="parenterrorlbl" CssClass="alert-danger" runat="server"></asp:Label>
											<asp:Label ClientIDMode="Static" ID="lblsaveerror" CssClass="alert-success" runat="server"></asp:Label>

											<asp:Repeater runat="server" ID="rptrdata" OnItemDataBound="rptrdata_ItemDataBound">
												<ItemTemplate>
													<p>
														<a name="Sect" class="section-heading"><font color="#000000" style="font-size: 20px;"><b><u>
															<asp:Label runat="server" ID="lblsection"></asp:Label></u></b></font></a>
													</p>
													<asp:Repeater runat="server" EnableViewState="true" ID="rptrchild" OnItemDataBound="rptrchild_ItemDataBound">
														<ItemTemplate>
															<asp:Label runat="server" ClientIDMode="Static" CssClass="validation-error" ID="lblerrormsg"></asp:Label>
															<a id="A4" name="A4"></a><span class="question-text"><font face="Arial">
																<br>
																<font style="font-size: 14px;">
																	<asp:Label ClientIDMode="Static" runat="server" ID="lblqorder"></asp:Label>
																</font></font><font style="font-size: 14px;"><font face="Arial" style="font-size: 14px;">
																	<asp:Label runat="server" ID="lblqtext"></asp:Label>


																	<img width="18" height="17" alt="" runat="server" class="question-img" visible="false" id="imghelptext" clientidmode="Static" src="images/question.png"></img>
															</span>


															<div class="response-set">
																<span class="litvalign">
																	<asp:Literal runat="server" ID="lblquescurrency" Visible="false"></asp:Literal>
																	&nbsp;
																</span>
																<asp:HiddenField ClientIDMode="Static" runat="server" ID="hidvalue" Value="" />
																<%-- <asp:TextBox enableviewstate="true" runat="server" clientidmode="Static" cols="62" rows="1" id="Q44" name="Q4"></asp:TextBox>--%>
																<%-- <div runat="server" id="litcontrol" clientidmode="Static"></div> --%>
																<textarea wrap="virtual" clientidmode="Static" runat="server" cols="62" rows="1" id="Q44" onkeypress="return validate(event,this)"></textarea>

																<asp:RadioButtonList runat="server" RepeatDirection="Horizontal" RepeatLayout="Table" ID="radgender" Visible="false">
																	<asp:ListItem Text="Male" Value="1"></asp:ListItem>
																	<asp:ListItem Text="Female" Value="2"></asp:ListItem>
																</asp:RadioButtonList>

																<asp:RadioButtonList runat="server" RepeatDirection="Horizontal" RepeatLayout="Table" ID="radOfficeManager" Visible="false">
																	<asp:ListItem Text="Yes" Value="1"></asp:ListItem>
																	<asp:ListItem Text="No" Value="2"></asp:ListItem>
																</asp:RadioButtonList>
															</div>
															<span class="text-block">
																<br>
															</span>
															<div class="response-set">

																<asp:Repeater runat="server" ID="rptrsubdata" OnItemDataBound="rptrsubdata_ItemDataBound">
																	<ItemTemplate>
																		<div class="lblposition1">
																			<b>
																				<asp:Label runat="server" ID="lblsub"></asp:Label></b>
																		</div>
																		<div class="clear"></div>
																		<div class="lblposition">
																			<asp:Label runat="server" ID="lblsubquest"></asp:Label>
																		</div>

																		<div class="txtposition">
																			<asp:Literal runat="server" ID="lblsubquescurrency" Visible="false"></asp:Literal>
																			&nbsp;
                                                                            <asp:TextBox ClientIDMode="Static" runat="server" ID="txtsubquestion" onkeypress="return validate1(event,this)"></asp:TextBox>


																		</div>


																		<asp:RadioButtonList runat="server" RepeatDirection="Horizontal" RepeatLayout="Table" ID="radsub" Visible="false">

																			<asp:ListItem Text="Yes" Value="1"></asp:ListItem>
																			<asp:ListItem Text="No" Value="2"></asp:ListItem>

																		</asp:RadioButtonList>

																		<div class="clear"></div>
																	</ItemTemplate>
																</asp:Repeater>

															</div>

														</ItemTemplate>
													</asp:Repeater>
												</ItemTemplate>
											</asp:Repeater>



											<span class="text-block">
												<br>
											</span>

											<!-- Do not delete or modify any of the hidden fields -->
											<input type="hidden" value="/Community/surveys/1069620006/16fca35f" id="PdcSurveyName" name="PdcSurveyName">
											<input type="hidden" value="/Community/surveys/1069620006/16fca35f002.htm" id="PdcCurrentPage" name="PdcCurrentPage">
											<input type="hidden" value="" id="PdcButtonPressed" name="PdcButtonPressed">
											<input type="hidden" value="3FC11B2616FCA35F08D23E822E58D8BB77" id="PdcSessionId" name="PdcSessionId">


											<p></p>
											<a name="END"></a>

											<asp:Button runat="server" ID="btnback" Text="Back" class="submit-button back" OnClick="Backpagebutton_Click" />
											<asp:Button runat="server" ID="btnNext" ClientIDMode="Static" Text="Next" class="submit-button next" OnClick="Nextpagebutton_Click" />
											<asp:Button runat="server" ID="btnSave" Text="Save" class="submit-button" OnClientClick="" OnClick="Savepagebutton_Click" />
											&nbsp;
 <asp:Button ClientIDMode="Static" runat="server" ID="btnSubmit" Text="Submit" class="submit-button" OnClick="Submitpagebutton_Click" Visible="false" />

											&nbsp;&nbsp;
                                            <asp:Button ClientIDMode="Static" runat="server" ID="btnPreview" Text="Preview" class="submit-button preview " OnClick="Previewbutton_Click" Visible="false" />

											<asp:Button runat="server" ID="btnbackbottom" Text="Back" class="submit-button back bottom" OnClick="Backpagebutton_Click" />
											<asp:Button runat="server" ID="btnNextbottom" ClientIDMode="Static" Text="Next" class="submit-button next bottom" OnClick="Nextpagebutton_Click" />


											<br>


											<asp:Button runat="server" ID="btnprint" Text="Print" class="submit-button" OnClick="btnprint_Click" />
											&nbsp;&nbsp;                                      
										</div>

										<div align="center" runat="server" id="progressbar1">
											<table runat="server" id="tblpaging" clientidmode="Static" summary="28%" class="progress-table">
												<%--<tbody><tr>
      <%--<td style="width:10%;" class="completed-cell">&nbsp;</td>
      <td style="width:10%;" class="completed-cell">&nbsp;</td>
      <td style="width:10%;" class="uncompleted-cell">&nbsp;</td>
      <td style="width:10%;" class="uncompleted-cell">&nbsp;</td>
      <td style="width:10%;" class="uncompleted-cell">&nbsp;</td>
      <td style="width:10%;" class="uncompleted-cell">&nbsp;</td>
      <td style="width:10%;" class="uncompleted-cell">&nbsp;</td>
      <td style="width:10%;" class="uncompleted-cell">&nbsp;</td>
      <td style="width:10%;" class="uncompleted-cell">&nbsp;</td>
      <td style="width:10%;" class="uncompleted-cell">&nbsp;</td>
      <td class="progress-text">28%</td>--%>
												<%--  </tr>
  </tbody>--%>
											</table>


											<div runat="server" id="divprogress" class="progress-text"></div>
										</div>


										<span class="perseus-link">
											<br>
										</span>
									</div>
								</div>
								<div id="bottom"><i></i></div>
							</td>
							<td class="v-MR"></td>
						</tr>
						<tr>
							<td class="v-BL">
								<div class="LeftPadding"></div>
							</td>
							<td class="v-BM">
								<div class="BottomPadding"></div>
							</td>
							<td class="v-BR">
								<div class="RightPadding"></div>
							</td>
						</tr>
					</tbody>
				</table>
				<div class="summary_center" tabindex="-1">
					<div class="summary_column last">
						<div class="summary_title">
							<strong>
								<a href="/en/Press/Pages/Default.aspx">Press</a>
							</strong>
						</div>
						<div class="summary_list">
							<ul>
								<li class="fisrt">
									<a href="/en/Press/News/Pages/Home.aspx">Press releases</a>
								</li>
								<li>
									<a href="/en/Press/LatestNews/Pages/home.aspx">News</a>
								</li>
								<li>
									<a href="/en/Press/Pages/Home.aspx">Media library</a>
								</li>
								<li>
									<a href="/en/Press/Pages/Videos.aspx">Videos</a>
								</li>
							</ul>
						</div>
					</div>
					<div class="summary_column">
						<div class="summary_title">
							<strong>
								<a href="/en/Investors/Pages/Home.aspx">Investors</a>
							</strong>
						</div>
						<div class="summary_list">
							<ul>
								<li class="fisrt">
									<a href="/en/Investors/KeyFigures/Pages/KeyFigures.aspx">Key Figures</a>
								</li>
								<li>
									<a href="/en/Investors/StockInformation/Pages/Home.aspx">Stock Information</a>
								</li>
								<li>
									<a href="/en/Investors/events/Pages/Events.aspx">Events</a>
								</li>
								<li>
									<a href="/en/Investors/Pages/PublicationsDownloads.aspx">Publications &amp; Downloads</a>
								</li>
								<li>
									<a href="/en/Investors/IndividualShareholderInformation/Pages/IndividualShareholderInformation.aspx">Individual Shareholder Information</a>
								</li>
								<li>
									<a href="/en/Investors/RegulatoryInformation/Pages/RegulatoryInformation.aspx">Regulatory Information</a>
								</li>
								<li>
									<a href="/en/Investors/Debt/Pages/default.aspx">Debt</a>
								</li>
								<li>
									<a href="/en/Investors/InvestorContacts/Pages/default.aspx">Investor Contacts</a>
								</li>
							</ul>
						</div>
					</div>
					<div class="summary_column">
						<div class="summary_title">
							<strong>
								<a href="/en/Talents/Pages/Home.aspx">Talents</a>
							</strong>
						</div>
						<div class="summary_list">
							<ul>
								<li class="fisrt">
									<a href="/en/Talents/OurPhilosophy/Pages/Our philosophy.aspx">Our philosophy</a>
								</li>
								<li>
									<a href="/en/Talents/Pages/OurValues.aspx">Our values</a>
								</li>
								<li>
									<a href="/en/Talents/CareerManagement/Pages/Home.aspx">Career management</a>
								</li>
								<li>
									<a href="/en/Talents/Jobs/Pages/Home.aspx">Jobs and profiles</a>
								</li>
								<li>
									<a href="/en/Talents/Pages/JoinEssilor.aspx">Join Essilor</a>
								</li>
							</ul>
						</div>
					</div>
					<div class="summary_column">
						<div class="summary_title">
							<strong>
								<a href="/en/BrandsAndProducts/Pages/Brandsproducts.aspx">Brands &amp; products</a>
							</strong>
						</div>
						<div class="summary_list">
							<ul>
								<li class="fisrt">
									<a href="/en/BrandsAndProducts/Lenses/Pages/Home.aspx">Lenses</a>
								</li>
								<li>
									<a href="/en/BrandsAndProducts/Non-prescription-lenses/Pages/home.aspx">Non-prescription lenses</a>
								</li>
								<li>
									<a href="/en/BrandsAndProducts/Instruments/Pages/Home.aspx">Instruments &amp; Equipment</a>
								</li>
							</ul>
						</div>
					</div>
					<div class="summary_column">
						<div class="summary_title">
							<strong>
								<a href="/en/EyeHealth/Pages/Home.aspx">Eye care</a>
							</strong>
						</div>
						<div class="summary_list">
							<ul>
								<li class="fisrt">
									<a href="/en/EyeHealth/UnderstandingVision/Pages/Home.aspx">Understanding vision</a>
								</li>
								<li>
									<a href="/en/EyeHealth/LensesForYourVision/Pages/Home.aspx">Lenses for your vision</a>
								</li>
								<li>
									<a href="/en/EyeHealth/VisionDefects/Pages/Home.aspx">Vision defects</a>
								</li>
								<li>
									<a href="/en/EyeHealth/WearersNeeds/Pages/default.aspx">Wearer's needs</a>
								</li>
								<li>
									<a href="/en/EyeHealth/DidYouKnow/Pages/default.aspx">Did You Know?</a>
								</li>
								<li>
									<a href="/en/EyeHealth/Pages/2012WorldSightDay.aspx">World Sight Day</a>
								</li>
							</ul>
						</div>
					</div>
					<div class="summary_column">
						<div class="summary_title">
							<strong>
								<a href="/en/Innovation/Pages/Home.aspx">Innovation</a>
							</strong>
						</div>
						<div class="summary_list">
							<ul>
								<li class="fisrt">
									<a href="/en/Innovation/InnovationByEssilor/Pages/Home.aspx">Innovation by Essilor</a>
								</li>
								<li>
									<a href="/en/Innovation/InnovationPortraits/Pages/Home.aspx">Innovation portraits</a>
								</li>
								<li>
									<a href="/en/Innovation/Pages/Partnerships.aspx">Partnerships</a>
								</li>
								<li>
									<a href="/en/Innovation/magazine/Pages/default.aspx">E-novation Magazine</a>
								</li>
							</ul>
						</div>
					</div>
					<div class="summary_column">
						<div class="summary_title">
							<strong>
								<a href="/en/Group/Pages/Home.aspx">Group</a>
							</strong>
						</div>
						<div class="summary_list">
							<ul>
								<li class="fisrt">
									<a href="/en/Group/OurStrategy/Pages/default.aspx">Our strategy</a>
								</li>
								<li>
									<a href="/en/Group/International/Pages/Home.aspx">International presence</a>
								</li>
								<li>
									<a href="/en/Group/Governance/Pages/Home.aspx">Governance</a>
								</li>
								<li>
									<a href="/en/Group/History/Pages/Home.aspx">History</a>
								</li>
								<li>
									<a href="/en/Group/Sustainable/Pages/Home.aspx">Sustainable enterprise</a>
								</li>
								<li>
									<a href="/en/Group/Shareholding/Pages/Home.aspx">Employee shareholding</a>
								</li>
								<li>
									<a href="/en/Group/Pages/EssilorPrinciples.aspx">Essilor Principles</a>
								</li>
							</ul>
						</div>
					</div>
					<div class="clear">
					</div>
				</div>


				<div id="footer_wrapper">
					<a class="foot_link" href="http://www.essilor.com/en/Pages/Contact.aspx">Contact</a>&nbsp;|&nbsp;<a class="foot_link" href="/en/Pages/OtherWebsites.aspx">Other Essilor Websites</a>&nbsp;|&nbsp;<a class="foot_link" href="http://www.essilor.com/en/Pages/Glossary.aspx">Glossary</a>&nbsp;|&nbsp;<a class="foot_link" href="http://www.essilor.com/en/Pages/Legal.aspx">Legal Notifications</a>&nbsp;|&nbsp;<a class="foot_link" href="http://www.essilor.com/en/Pages/SiteMap.aspx">Site map</a>
				</div>
			</div>
		</div>
		<style>
			#footer_wrapper {
				font-size: 11px;
				text-align: center;
			}

				#footer_wrapper a {
					color: #8e9699;
					text-decoration: none;
				}

					#footer_wrapper a:hover {
						color: #ab162b;
					}
		</style>
	</form>
</body>
</html>
