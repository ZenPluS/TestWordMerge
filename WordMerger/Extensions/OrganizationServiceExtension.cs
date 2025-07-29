using Microsoft.Xrm.Sdk;

namespace WordMerge.Extensions
{
    internal static class OrganizationServiceExtension
    {
        internal static TOrganizationResponse Execute<TOrganizationResponse, TOrganizationRequest>(
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
