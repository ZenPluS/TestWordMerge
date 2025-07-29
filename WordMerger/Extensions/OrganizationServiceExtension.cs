using Microsoft.Xrm.Sdk;

namespace WordMerge.Extensions
{
    public static class OrganizationServiceExtension
    {
        public static TOrganizationResponse Execute<TOrganizationResponse, TOrganizationRequest>(
            this IOrganizationService service,
            TOrganizationRequest request
            )
            where TOrganizationRequest: OrganizationRequest
            where TOrganizationResponse : OrganizationResponse
        {
            return (TOrganizationResponse)service.Execute(request);
        }
    }
}
