using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;

namespace TestWordMerge.Extensions
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
