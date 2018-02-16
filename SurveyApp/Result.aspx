<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Result.aspx.cs" Inherits="SurveyApp.Result" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
	<link href="Styles/axure_rp_page.css" rel="stylesheet" type="text/css" />
	<link href="Styles/axurerp_privacy_notice.css" rel="stylesheet" type="text/css" />
	<link href="Styles/jquery-ui-themes.css" rel="stylesheet" type="text/css" />
	<link href="Styles/StyleSheet1.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="Scripts/jquery-1.4.1.js"></script>
	<script src="Scripts/jquery-1.4.1.min.js" type="text/javascript"></script>
	<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
	<style>
		.lblposition {
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

				   if (!(keycode == 8 || keycode == 46) && (keycode < 48 || keycode > 57)) {
					   return false;
				   }
				   else {
					   //Condition to check textbox contains ten numbers or not
					   if (phn.value.length < 10) {
						   return true;
					   }
					   else {
						   return false;
					   }
				   }
			   }

			   else {
				   return true;
			   }


		   }

	</script>
	<script type="text/javascript">


		   function validate1(key, obj) {

			   var keycode = (key.which) ? key.which : key.keyCode;
			   var phn = document.getElementById('txtsubquestion');



			   if (!(keycode == 8 || keycode == 46) && (keycode < 48 || keycode > 57)) {
				   return false;
			   }
			   else {
				   //Condition to check textbox contains ten numbers or not
				   if (phn.value.length < 10) {
					   return true;
				   }
				   else {
					   return false;
				   }
			   }
		   }

       

	</script>
	<script type="text/javascript">


		   $(document).ready(function () {


			   $("#btnNext").click(function (e) {


				   //alert('t');
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
								   //return false;
							   }

							   else {

								   document.getElementById("lblerror" + controlid).innerHTML = "";

							   }

						   }

					   }

				   });

			   });
		   });


	</script>
	<title></title>
</head>
<body>
	<form id="form1" runat="server">
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
				<table cellspacing="0" cellpadding="0" border="0" class="v-Table">
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
								<div id="surveyBlock">
									<div id="surveyBlockNest">

										<div runat="server" id="firstpage" clientidmode="Static" class="mydiv">

											<asp:Label ID="lblmsg" runat="server"></asp:Label>
											<a></a>
											<a></a>
											<br />
											<div style="text-align: center;">
												<asp:Button runat="server" ID="btnHome" Text="Back to Home Page" class="submit-buttonhome" OnClick="Homepagebutton_Click" />

											</div>
											<div style="text-align: center;">
												<asp:Button runat="server" ID="btnDownload" Text="Download Reports" class="submit-buttonhome" OnClick="btnDownload_Click" />

											</div>
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
				<div class="summary_center">
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
			</div>
		</div>
	</form>
</body>
</html>
