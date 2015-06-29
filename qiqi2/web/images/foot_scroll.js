var foot_timer;
function foot_scroll(){
	$(".foot_scroll_ul").animate({top:"-70px"},700,function(){
			$(".foot_scroll_ul li:first").appendTo(".foot_scroll_ul");
			$(".foot_scroll_ul").css("top","0");
	})
	foot_timer=window.setTimeout("foot_scroll()",10000);
}
foot_timer=window.setTimeout("foot_scroll()",10000);

document.writeln("<link href=\"images\/foot_croll.v2.css\" type=\"text\/css\" rel=\"stylesheet\" \/>");

document.writeln("<div class=\"foot_scroll\">");
document.writeln("	<ul class=\"foot_scroll_ul\">");
document.writeln("		<li class=\"b1\"><span>本站互动平台是一家专业提升店铺信誉的平台，汇聚上万名网商，而且每天还在以数千个的速度在不断递增。");
document.writeln("网店无排名无生意不用担心，来吧，我们给您满意的解决方案。<\/span><\/li>");
document.writeln("		<li class=\"b2\"><span>安全是我们的保证，免费是我们的承诺。只要您按平台的流程操作永久免费，全国各地上万名会员在这里免费互刷，您还在犹豫吗？");
document.writeln("您在这里可以学的开店的经验，与钻石皇冠卖家共同分享心得，让您店铺生意一升在升。<\/span><\/li>");
document.writeln("		<li class=\"b3\"><span>平台的宗旨是大家帮助大家，一份耕耘一份收获，帮助别人就是在提升自己，接任务赚发布点，发任务消耗发布点");
document.writeln("官方永久回收发布点0.25元每个，如果您觉的您的时间充足，我们给你提供日赚百元的兼职网赚事业。<\/span><\/li>");
document.writeln("		<li class=\"b4\"><span>立即免费注册开始行动吧！本站提供安全的资金担保服务，让您操作零风险，并提供随时提现的解决方案");
document.writeln("即使你在做任务的过程中出现了纠纷，本站也将提供公正的争议仲裁服务，保证您的资金安全。<\/span><\/li>");
document.writeln("		<li class=\"b5\"><span>本站是全国同类网站中最大的平台之一。我们的一切努力都是为了信誉，您的满意是我们的服务遵旨");
document.writeln("让你自由而快乐地在线提升店铺排名，打造火爆生意的同时还能赚钱；让失落的网商有了希望，让失业的朋友有了赚钱机会。<\/span><\/li>");
document.writeln("		<li class=\"b6\"><span>请将本站加入你的收藏夹，以免以后你需求的时候忘记我们。如果你是自由工作者，你可以在这里挣到");
document.writeln("实现自己的梦想。如果你是店长，你可以在这里打造您的金冠店铺，更低的成本，更多的选择。<\/span><\/li>");
document.writeln("	<\/ul> ");
document.writeln("<\/div>")
