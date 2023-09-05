using EWS.NetEnhanced.Data;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

namespace EWS.NetEnhanced
{
    public class EwsCallerService
    {
        private readonly ExchangeService _service;
        private readonly string _accessToken;

        public EwsCallerService(string appId, string clientSecret, string tenantId, string[] scopes, string url, string email)
        {
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

            _service = new ExchangeService();

            var cca = ConfidentialClientApplicationBuilder
               .Create(appId)
               .WithClientSecret(clientSecret)
               .WithTenantId(tenantId)
               .Build();

            var clientTask = System.Threading.Tasks.Task.Run(() => cca.AcquireTokenForClient(scopes).ExecuteAsync());
            clientTask.Wait();

            _accessToken = clientTask.Result.AccessToken;
            _service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, email);

            //Include x-anchormailbox header
            _service.HttpHeaders.Add("X-AnchorMailbox", email);
            _service.Credentials = new OAuthCredentials(_accessToken);
            _service.Url = new Uri(url);
        }

        public FindItemsResults<Item> GetEmails(DateTime since, int pageSize, bool? hasAttachments = null)
        {
            return _service.FindItems(WellKnownFolderName.Inbox, SetFilter(since, null, hasAttachments), CreateView(pageSize));
        }

        public FindItemsResults<Item> GetEmails(DateTime since, string subject, int pageSize, bool? hasAttachments = null)
        {
            return _service.FindItems(WellKnownFolderName.Inbox, SetFilter(since, subject, hasAttachments), CreateView(pageSize));
        }

        public string AccessToken()
        {
            return _accessToken;
        }

        public Uri GetUrl()
        {
            return _service.Url;
        }

        public FindItemsResults<Item> Search(SearchFilter filter, int pageSize)
        {
            return _service.FindItems(WellKnownFolderName.Inbox, filter, CreateView(pageSize));
        }

        public ServiceResponseCollection<GetAttachmentResponse> GetAttachments(string[] attachmentIds, BodyType? bodyType, IEnumerable<PropertyDefinitionBase> additionalProperties)
        {
            return _service.GetAttachments(attachmentIds, bodyType, additionalProperties);
        }

        public ServiceResponseCollection<GetAttachmentResponse> GetAttachments(Attachment[] attachments, BodyType? bodyType, IEnumerable<PropertyDefinitionBase> additionalProperties)
        {
            return _service.GetAttachments(attachments, bodyType, additionalProperties);
        }

        public AttachmentCollection GetAttachments(ItemId id)
        {
            EmailMessage message = EmailMessage.Bind(_service, id, new PropertySet(ItemSchema.Attachments));

            return message.Attachments;
        }

        public void Send(string subject, string body, string to, bool isHtml = false)
        {
            EmailMessage message = new EmailMessage(_service)
            {
                Subject = subject,
                Body = new MessageBody { Text = body, BodyType = isHtml ? BodyType.HTML : BodyType.Text }
            };

            message.ToRecipients.Add(to);
            message.Send();
        }

        private static SearchFilter SetFilter(DateTime startDate, string? subject = null, bool? hasAttachments = null)
        {
            SearchFilter dateFilter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, startDate);

            SearchFilter.SearchFilterCollection coll = new SearchFilter.SearchFilterCollection(LogicalOperator.And)
            {
                dateFilter
            };

            if (!string.IsNullOrEmpty(subject))
                coll.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, subject));

            if (hasAttachments != null)
                coll.Add(new SearchFilter.IsEqualTo(ItemSchema.HasAttachments, (bool)hasAttachments));

            return coll;
        }

        private bool CertificateValidationCallBack(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            if (sslPolicyErrors == SslPolicyErrors.None)
                return true;

            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if (chain != null && chain.ChainStatus != null)
                {
                    foreach (X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                           (status.Status == X509ChainStatusFlags.UntrustedRoot))
                            continue;

                        if (status.Status != X509ChainStatusFlags.NoError)
                            return false;
                    }
                }

                return true;
            }
            else
            {
                return false;
            }
        }

        private static ItemView CreateView(int pageSize)
        {
            ItemView view = new ItemView(pageSize)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.DateTimeReceived, ItemSchema.HasAttachments)
            };

            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

            return view;
        }
    }
}
