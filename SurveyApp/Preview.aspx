<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Preview.aspx.cs" Inherits="SurveyApp.Preview" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
	Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
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
											<asp:Label ClientIDMode="Static" ID="lblsaveerror" CssClass="alert-danger" runat="server"></asp:Label>

											<asp:Repeater runat="server" ID="rptrdata" OnItemDataBound="rptrdata_ItemDataBound">
												<ItemTemplate>
													<p><a name="Sect" class="section-heading"><font color="#000000" style="font-size: 20px;"><b><u>
														<asp:Label runat="server" ID="lblsection"></asp:Label></u></b></font></a></p>
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

																<asp:Label Style="color: #000; font-weight: bold; padding: 0 0 0 15px;" ClientIDMode="Static" runat="server" cols="62" rows="1" ID="Q44" Enabled="false" onkeypress="return validate(event,this)"></asp:Label>

																<asp:RadioButtonList runat="server" RepeatDirection="Horizontal" RepeatLayout="Table" ID="radgender" Visible="false">
																	<asp:ListItem Text="Male" Value="1" Enabled="false"></asp:ListItem>
																	<asp:ListItem Text="Female" Value="2" Enabled="false"></asp:ListItem>
																</asp:RadioButtonList>

																<asp:RadioButtonList runat="server" RepeatDirection="Horizontal" RepeatLayout="Table" ID="radOfficeManager" Visible="false">
																	<asp:ListItem Text="Yes" Value="1" Enabled="false"></asp:ListItem>
																	<asp:ListItem Text="No" Value="2" Enabled="false"></asp:ListItem>
																</asp:RadioButtonList>
															</div>
															<span class="text-block"></span>
															<div class="response-set">

																<asp:Repeater runat="server" ID="rptrsubdata" OnItemDataBound="rptrsubdata_ItemDataBound">
																	<ItemTemplate>
																		<div class="lblposition1"><b>
																			<asp:Label runat="server" ID="lblsub"></asp:Label></b></div>
																		<div class="clear"></div>
																		<div class="lblposition">
																			<asp:Label runat="server" ID="lblsubquest"></asp:Label></div>

																		<div class="txtposition">
																			<asp:Literal runat="server" ID="lblsubquescurrency" Visible="false"></asp:Literal>
																			&nbsp;
																			<asp:Label ClientIDMode="Static" Style="color: #000; font-weight: bold;" runat="server" ID="txtsubquestion" onkeypress="return validate1(event,this)" Enabled="false"></asp:Label>


																		</div>


																		<asp:RadioButtonList runat="server" RepeatDirection="Horizontal" RepeatLayout="Table" ID="radsub" Visible="false">

																			<asp:ListItem Text="Yes" Value="1" Enabled="false"></asp:ListItem>
																			<asp:ListItem Text="No" Value="2" Enabled="false"></asp:ListItem>

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


											<asp:Button ClientIDMode="Static" runat="server" ID="btnSubmit" Text="Submit" class="submit-button" OnClick="Submitpagebutton_Click" Visible="true" />
											<asp:Button ClientIDMode="Static" runat="server" ID="btnback" Text="Back to Home Page" class="submit-button" OnClick="Backpagebutton_Click" Visible="true" />

											<%-- <asp:Label ID="lbl1year" runat="server" Visible="false" ></asp:Label>
 <asp:Label ID="lblname11" runat="server" Visible="false" ></asp:Label>
<asp:Label ID="lblpractice11" runat="server" Visible="false" ></asp:Label>

                                             <asp:Label ID="lblyeartext1" runat="server" Visible="false" ></asp:Label>--%>


											<br>
										</div>


										<%-- <asp:Button runat="server" ID="btnprint" Text="Print" class="submit-button" OnClick="btnprint_Click"/>                                          --%>
										<%--<div align="center" runat="server" id="progressbar1">
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

										<%--  </table>

  
  <div runat="server" id="divprogress" class="progress-text"></div>
</div>--%>


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
