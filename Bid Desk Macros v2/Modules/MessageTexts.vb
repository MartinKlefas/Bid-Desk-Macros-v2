Module MessageTexts
    Public Const DRExpire As String = "<br>&nbsp; Deal Registration <strong>%dealID%</strong> for customer: <strong>%customer%</strong> will expire shortly.<br>&nbsp;  Please could you Let Me know If you'd like me to renew it, or if the opportunity is dead?<br> Thanks,<br> Martin."
    Public Const drloglink As String = "<br><a href=""https://insightonlinegbr-my.sharepoint.com/personal/martin_klefas_insight_com/Documents/New%20DR%20Log.xlsx?web=1"" >Click here for an automatically updated deal status report(which you should be able to filter)</a>"
    Public Const drDecision As String = "<br>&nbsp; Please see below/attached the vendor's decision on your deal registration.<br> Thanks,<br> Martin."
    Public Const dellDecline As String = "<br>&nbsp; Please note that a declined Dell DR can always be appealed if you feel you have a strong reason as to why it should have been approved - please contact Rebecca.Pas@dell.com & shireen_kandola@dell.com to make your case.<br>"
    Public Const sqFwdMessage As String = "<br>&nbsp; Please see attached pricing from distribution.<br> Thanks,<br> Martin."
    Public Const opgFwdMessage As String = "<br>&nbsp; Please see attached the OPG pricing from distribution, you can now create a UPR creation ticket, but no longer need to attach this to it.<br> Thanks,<br> Martin."
    Public Const HPPublishMessage As String = "<br>&nbsp; Your SmartQuote has been approved by HP, and is now being priced by distribution. We should have the first responses back very shortly. "
    Public Const HPExtensionSubmitted As String = "<br>&nbsp; I've submitted your request for an extension in the portal, you should hopefully receive an updated response shortly. Please note though that pricing may be different on the renewed deal. <br> Thanks,<br> Martin."
    Public Const DellExtensionSubmitted As String = "<br>&nbsp; I've submitted your request for an extension in the portal, and it has been extended by 90 days. Please ensure the continued validity of any quotes before passing them on to your customer. <br> Thanks,<br> Martin."
    Public Const WonMessage As String = "<br>&nbsp; As below, your SQ has been set to won. Please wait for confirmation of success from the vendor before creating your UPR ticket.<br> Thanks,<br> Martin."
    Public Const DeadMessage As String = "<br>&nbsp; Thank you for letting me know, I've updated my records.<br> Thanks,<br> Martin."
    Public Const SubmitMessage As String = "<br>&nbsp;I've created the below for you with %VENDOR% (ref: %DEALID%). %NDT%<br>&nbsp;Please check that everything is correct and let me know asap if there are any errors.<br> Regards, Martin."

    Public Const CiscoAttachComment As String = "Please See attached the vendor quote in dollars. Pre-Sales team: Please move this ticket to Neil or Duncan for an Internal Cost Document and Customer facing quote to be created."

    Public Const TicketSubmitMessage As String = "I've created %DEALID% for you on the %VENDOR% portal.
        %VENDOR% will endeavor to turn this around for you as soon as possible, but if there are any unexpected delays then please reach out and we can escalate it after the normal SLA period. Regards, Martin."



    Public Const NDTCreateMessage As String = "For internal tracking purposes only, I have also created NextDesk ticket #%NDT%"

    Public Const NDTUseMessage As String = "I have also updated NextDesk Ticket #%NDT% with this information for consistency of internal messaging, and easier tracking."

    Public Const NoNDTMessage As String = "No NextDesk ticket was created for this action."

    Public Const CloneLaterMessage As String = "Due to the age of this deal registration, it can no longer be extended. In order for our protection to continue, we must now clone it onto a fresh deal registration.<br> If this cloning takes place before our existing deal registration ends, the new one will be automatically rejected as a duplicate of the old, and so this will be cloned for you on %CLONEDATE%. I have set myself a reminder to perform this clone, and so no further action is required from you.<br><br>"

    Public Const HolidayMessage As String = "<p><span style=""color: #ff0000;""><strong>Please Note:&nbsp;</strong>This email has been sent automatically while I am out of the office. </span>Replying to it will not elicit a response or any action on my return - please follow the instructions in my out of office message to ensure that your request is processed.</p><p>It should be noted that any emails received in reply to this message will be deleted on my return, and no action will be taken. Your email has not been forwarded, and has not been attached to the relevant nextdesk ticket.&nbsp;</p><p>If you need to get hold of me while I am away, your line manager has been left with instructions on how and when doing so is more appropriate than escalating your request through the normal channels.</p>"

    Public Const CloneTicketMessage As String = "This deal will be cloned at a later date, at which point a new ticket will be raised"

    Public Const MoreInfoRequested As String = "Please see %BELOW% a request for more information from %VENDOR% regarding deal ref: %DEALID% / NDT#: %NDT%. As outlined in the request, %VENDOR% requires this information within 4 business days or request will be denied.<br><br>If you decide not to proceed with this request before this deadline passes, please could you let me know by email as soon as possible."

    Public Const VendorInfoUpdate As String = "Please see %BELOW% an update from %VENDOR% regarding deal ref: %DEALID% / NDT#: %NDT%."


    Public Const EmailForwardMessage As String = "Please see %BELOW% an email from %VENDOR% regarding deal ref: %DEALID% / NDT#: %NDT%. Please can you reach out directly to %VENDOR% to discuss this opportunity, copying me in only if there is a need for me to make changes in the %VENDOR% portal."


    Public Const PreSubMoreInfo As String = "Microsoft has asked that we ensure that we are submitting deal registrations with more detailed information on than that which has been submitted on this ticket. This reduces the number of rejections that we will receive and decreases the overall time to approval. It also lessens both your workload and that of our valued partner account managers at Microsoft." & vbCrLf & vbCrLf & vbCrLf & "Please could you ensure that you have included ALL of the following information in your ticket and if not add it as a log so that it can be included in the initial submission: " & vbCrLf & "How will the end customer use these units?" & vbCrLf & "How did you come across this opportunity?" & vbCrLf & "Please also provide information regarding the pre-sales activity you have performed within the past 3 months to cultivate this opportunity for Surface. Pre-sales activity is any communication between the reseller and the customer that show Surface is being discussed." & vbCrLf & "• Please provide copies or descriptions of communication between you and the customer, this could include: meetings, phone calls, emails, and/or demos that were performed/provided." & vbCrLf & "• Copies of emails showing Pre-Sales Activity can be attached to the deal registration" & vbCrLf & "• Descriptions of Pre-Sales Activity must include dates of meetings/phone calls, etc."

    Public Const CiscoDRType As String = "Hi, in your ticket request you've asked for Deal Registration for which there are two alternatives from Cisco:

1) Hunting: This is for situations in which you have pro-actively identified an opportunity and are bringing it to Cisco. This could be for customers who have previously used other vendors, or alternatively are trying a new portion of the Cisco portfolio.

2) Teaming: This is for opportunities on which you're collaborating with Cisco, having landed and expanded a customer into new opportunities this is intended to reward the joint work that you have put in with a named Cisco Account Manager.

I have attached the two alternative questionnaires that apply to these two scenarios to the ticket, and will apply for pricing based on your response.

There are further discounts available based on things like our pre-sales (time) investments, so please reach out if neither of the above scenarios applies and we'll work out which discounts are most appropriate for your opportunity."


    Public LabelMessages As New Dictionary(Of String, String) From {
        {"Login", "Logging into CCW..."},
        {"NewDeal1", "Creating Deal (page 1) ..."},
        {"NewDeal2", "Creating Deal (page 2) ..."},
        {"NewDeal3", "Creating Deal (page 3) ..."},
        {"NewDeal4", "Creating Deal (page 4) ..."},
        {"DL1", "Finding the deal..."},
        {"DL2", "Exporting Quote..."},
        {"DL3", "Waiting for file download..."},
        {"AM1", "Looking up AM Details..."},
        {"LenovoLogin", "Logging into LBP..."},
        {"Searching", "Searching for the bid"},
        {"Sending", "Sending the bid to westcoast"}
        }

    Public Const CiscoAMFail As String = "The Cisco portal did not yet show the Account Manager. While there is no requirement for you to do so, discussing this deal with them will more than likely speed up the approvals process, and decrease your chances of it being rejected or ignored. If you would like me to check again shortly so that you can reach out to the right person, please let me know and I will."

    Public Const CiscoAMMessage As String = "The Cisco portal shows the Account Manager to be %AM%. While there is no requirement for you to do so, discussing this deal with them will more than likely speed up the approvals process, and decrease your chances of it being rejected or ignored."

    Public Const CiscoAMTeamMessage As String = "The Cisco portal shows that this account is managed by a team as below:
%AM%

While there is no requirement for you to do so, discussing this deal with them will more than likely speed up the approvals process, and decrease your chances of it being rejected or ignored."


    Public Const FindAMMessage As String = "While there is no requirement for you to do so, discussing this deal with the vendor Account Manager will more than likely speed up the approvals process, and decrease your chances of it being rejected or ignored.

 Please let me know if you'd like me to find out who the Vendor Account Manager for this customer is."

End Module
