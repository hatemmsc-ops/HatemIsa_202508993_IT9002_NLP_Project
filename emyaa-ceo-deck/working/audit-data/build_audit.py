import csv
H=['id','source_slide','displayed_figure','metric','numerator','denominator','comparison_period','recalculated','rounding_check','wording_check','source','status','correction_required']
def pct(a,b): return f"{a/b*100:.2f}%"
R=[
['P01','6','793%','Active users, week over week','384 (21-27 Sep)','~43 (14-20 Sep, implied; not visible)','Last 7 days vs previous 7 days',f"{(384-43)/43*100:.1f}% if prior = 43",'OK only if prior = 43','Overstated: slide says "Following the Saudi test launch". GA card covers all sources and all countries; the Saudi figure on the screenshot is the realtime 30-minute panel (7 users), not the 7-day total. Paid testing began 22 Sep, inside the window, but attribution is not shown.','GA home card screenshot (Enhanced slide 6 image)','Unverified','Label "Reported increase; source calculation pending verification". Show prior-period value from a GA export. Drop causal wording.'],
['P02','screenshot only','741.9%','Event count, week over week','2.5K (rounded)','~297 (implied)','Last 7 days vs previous 7 days','Cannot recalculate (2.5K is rounded)','n/a','Not shown in deck. Note: the line chart on the screenshot plots Event count, not Active users.','GA home card screenshot','Unverified','Do not use without export.'],
['P03','2, 3, 6','69.7%','Share of website sessions from Paid (Google Ads)','409 (derived)','587 sessions','Since 22 Sep (end date not stated)',pct(409,587),'OK','"Already 69.7% of website traffic" is fair only for the post-launch window. LinkedIn paid traffic is not split out.','Enhanced deck; no GA export supplied','Unverified','Keep only as "reported"; add date range end; supply GA source/medium export.'],
['P04','6','15.7%','Share of sessions, Direct','92 (derived)','587','Since 22 Sep',pct(92,587),'OK','None','Enhanced deck','Unverified','Needs GA export.'],
['P05','6','6.6%','Share of sessions, Organic search','39 (derived)','587','Since 22 Sep',pct(39,587),'OK','None','Enhanced deck','Unverified','Needs GA export.'],
['P06','6','6.6%','Share of sessions, Referral','39 (derived)','587','Since 22 Sep',pct(39,587),'OK','None','Enhanced deck','Unverified','Needs GA export.'],
['P07','6','98.6%','Top four sources as share of sessions','579','587','Since 22 Sep',pct(579,587),'OK','None','Enhanced deck','Verified (arithmetic)','None. Underlying counts still need export.'],
['P08','7','54.5%','English share of ad sessions','212','388 (EN+AR) or 389',"Since 23 Sep",f"{pct(212,388)} on 388; {pct(212,389)} on 389",'Displayed pair sums to 99.7%','Denominator not stated; displayed value only matches a 389 total (1 session outside EN/AR).','Enhanced deck','Unverified','Show counts (212 vs 176) and drop the share, or state the denominator.'],
['P09','7','45.2%','Arabic share of ad sessions','176','388 or 389','Since 23 Sep',f"{pct(176,388)} on 388; {pct(176,389)} on 389",'See P08','See P08','Enhanced deck','Unverified','As P08.'],
['P10','7','54.3%','English share of active users','208','383','Since 23 Sep',pct(208,383),'OK','None','Enhanced deck','Verified (arithmetic)','Source export still needed.'],
['P11','7','45.7%','Arabic share of active users','175','383','Since 23 Sep',pct(175,383),'OK','None','Enhanced deck','Verified (arithmetic)','Source export still needed.'],
['P12','7','57.4%','English share of engaged sessions','206','359','Since 23 Sep',pct(206,359),'OK','None','Enhanced deck','Verified (arithmetic)','Source export still needed.'],
['P13','7','42.6%','Arabic share of engaged sessions','153','359','Since 23 Sep',pct(153,359),'OK','None','Enhanced deck','Verified (arithmetic)','Source export still needed.'],
['P14','2, 3, 7','97.17% (97.2%, 97%)','English engagement rate','206 engaged sessions','212 sessions','Since 23 Sep',pct(206,212),'OK','"Leads on every metric" is accurate for the six metrics shown. Early data; may reflect creative, not audience.','Enhanced deck','Verified (arithmetic)','Keep with "since 23 Sep". Note launch date is 22 Sep: explain 1-day offset.'],
['P15','2, 3, 7','86.93% (86.9%, 87%)','Arabic engagement rate','153','176','Since 23 Sep',pct(153,176),'OK','As P14','Enhanced deck','Verified (arithmetic)','As P14.'],
['P16','new','10.2 pts','EN minus AR engagement rate','97.17%','86.93%','Since 23 Sep','10.24 percentage points','OK','Use "points", not "%".','Derived from P14, P15','Verified (arithmetic)','n/a'],
['P17','2','+6,075%','LinkedIn follower growth','531 - 284 = 247','284','Last 30 days',pct(247,284)+' (follower growth)','n/a','Wrong as worded. +6,075% only works as 247 new followers vs 4 in the prior 30 days, i.e. growth in new-follower adds, not in followers.','Enhanced deck slide 2 and slide 4 chart','Incorrect','Replace with "+247 followers in 30 days (+87%)", still unverified until export. Also resolve 586 to 531 drop on slide 8.'],
['P18','4','+118%','LinkedIn impressions','559','~256 (implied)','13-27 Sep vs prior period',f"{(559-256)/256*100:.1f}% if prior = 256",'OK if prior = 256','Slide 2 calls this "double digits"; it is triple digits.','Enhanced deck; no LinkedIn export','Unverified','Supply export; fix "double digits" wording.'],
['P19','4','+155%','LinkedIn reactions','28','11 (implied)','13-27 Sep vs prior period',f"{(28-11)/11*100:.1f}%",'OK (154.5% rounds to 155%)','As P18','Enhanced deck; no LinkedIn export','Unverified','Supply export.'],
['P20','2, 4','73.3% / 25.9% / 0.7%','Instagram share of views by format (Reels / Stories / Posts)','n/a','2,271,132 views','Labelled "all-time"','Sum 99.9%','OK (rounding)','Period label conflicts with evidence: one reel (Musheera) shows 5.3M views, more than the 2.27M "all-time" total. The export likely covers a limited period.','Enhanced deck; no IG export','Unverified','Confirm export date range and relabel.'],
['P21','2, 4','89.5% / 3.8% / 6.8%','Instagram share of interactions by format','n/a','8,311 interactions','Labelled "all-time"','Sum 100.1%','OK (rounding)','As P20','Enhanced deck; no IG export','Unverified','As P20.'],
['P22','new','+15.9%','App downloads vs prior pull','480 - 414 = 66','414','"Prior pull", dates not stated',pct(66,414),'n/a','Not shown in Enhanced deck; only usable once both pull dates are known.','Enhanced deck slide 3','Unverified','Do not publish until pull dates are known.'],
]
with open('output/audit/Percentage-Audit.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(H); w.writerows(R)

T=['id','claim_or_figure','source_slide','evidence_found','evidence_file','status','note']
S=[
['T01','Instagram 2,116 followers','2,3','None beyond deck text','-','Unverified','Needs IG export'],
['T02','Instagram 2,271,132 views, 8,311 interactions ("all-time")','4','Deck table only','-','Unverified','Conflicts with 5.3M single-reel views (see P20)'],
['T03','LinkedIn 531 followers; 284 thirty days ago','3,4','Deck chart only','-','Unverified','Conflicts with "dropped 586 to 531" on slide 8'],
['T04','LinkedIn impressions 559, reactions 28 (13-27 Sep)','4','Deck table only','-','Unverified','Needs LinkedIn export'],
['T05','587 website sessions since 22 Sep','3,6','Deck only','-','Unverified','Needs GA export'],
['T06','558 new users since campaign launch','6','Deck only','-','Unverified','Higher than 384 active users on the 21-27 Sep GA card; windows differ, need export'],
['T07','384 active users, 21-27 Sep, +793%','6','GA home card screenshot','input/analytics/GA-home-card-21-27Sep-extracted-from-Enhanced-slide6.png','Figure verified on screenshot; % calc and attribution unverified','Screenshot legible'],
['T08','480 app downloads vs 414 prior pull','3','Deck only','-','Unverified','Pull dates missing'],
['T09','Paid media: Google Search to Starter/Emerging; LinkedIn lead gen to Affluent','6','Deck text only','-','Unverified','Needs Ignite media plan'],
['T10','$30,000 budget Sep to Dec, managed by Ignite, invoice TI7541','6','Deck text only','-','Unverified','TI7541 document not supplied; deck says "invoice", brief says "budget or ticket". TI7542 referenced in brief, not found anywhere'],
['T11','Paid testing began 22 Sep; Saudi residents only','2,6','Deck text only','-','Unverified','GA line rises from 24-25 Sep'],
['T12','EN vs AR ad metrics since 23 Sep','7','Deck table only','-','Arithmetic verified; source unverified','Needs GA or Google Ads export'],
['T13','Mention & Win: 23 Sep to 20 Oct, timed before salary days, GCC-wide','5,9','Deck text only; no T&Cs','-','Unverified','Mechanics, eligibility and prize terms not in review set'],
['T14','Mention & Win $500 post: 186K views','10','Instagram grid screenshot','input/campaign-evidence/Instagram-grid-extracted-from-Enhanced-slide10.png','Verified (screenshot)','Capture date unknown'],
['T15','Mention & Win launch reel: 26.2K views, 1,129 likes','5','Screenshot shows 26.3K on "Open your account, chance to win $500" post; likes not shown','Instagram grid screenshot','Views outdated (26.3K later); likes unverified','Confirm this post belongs to Mention & Win'],
['T16','$5,000 trading campaign','brief','Most Active Trader Raffle T&Cs: $500 to $5,000 prize by volume tier, 1 Aug to 31 Dec 2026','input/campaign-evidence/Most-Active-Trader-Raffle-TandC-EN-excerpt.md','Verified as separate campaign (T&Cs only)','Live status not confirmed'],
['T17','Ali Sabeel reel 1.3M views; "Completed 18 Sep"','5,10','1.3M on screenshot (matched by view count); completion inferred from last Ignite invoice 18 Sep','Instagram grid screenshot','Views verified; completion unverified','Rollout row says Aug posted; label as Posted'],
['T18','Musheera reel 5.3M views','10','Screenshot 5.3M','Instagram grid screenshot','Verified (screenshot)','-'],
['T19','How-to video (human presenter) 319 views, 28 interactions','5','Screenshot tile shows 332 views (likely later capture)','Instagram grid screenshot','Views outdated; interactions unverified','Tile match by view count only'],
['T20','Abdulelah Al Harbi video posted (Sep)','5','Deck text only','-','Unverified','No link or metric'],
['T21','Manal Talal in production, mid-October','5,9','Deck text only','-','Unverified','Needs brief/production tracker'],
['T22','Hamed Al Bloushi scheduled November','5,9','Deck text only','-','Unverified','Needs content calendar'],
['T23','Saudi Market Size infographic 515 views','10','No matching tile (grid shows 262, 347, 549 on graphics)','Instagram grid screenshot','Unverified','-'],
['T24','Gross deposits $12,206 confirmed by Malek','2,8','Verbal confirmation only per deck; period not stated','-','Unverified','Needs Malek export'],
['T25','LinkedIn follower drop 586 to 531','8','Deck only','-','Unverified','Contradicts growth story'],
['T26','"Presenter-led reels outperform static graphics by orders of magnitude"','10','Screenshot shows $500 static graphic at 186K and presenter how-to at 332','Instagram grid screenshot','Incorrect as a general claim','Top two posts are presenter-led; boost spend likely differs'],
['T27','X and TikTok performance','brief','Not present in any supplied file','-','Missing','Pending exports (slides 15-18)'],
['T28','Market Pulse data','brief','Not present in any supplied file','-','Missing','-'],
]
with open('output/audit/Source-Traceability.csv','w',newline='') as f:
    w=csv.writer(f); w.writerow(T); w.writerows(S)
print(len(R),len(S))
