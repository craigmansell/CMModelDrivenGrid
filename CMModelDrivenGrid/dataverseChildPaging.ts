import { childLogicalName, ChildFetchResult } from './childRecords';

/** Use the authenticated app's same-origin public Dataverse endpoint. The PCF
 * client wrapper can repeat FetchXML page 1; HTTP preserves page/cookie values. */
export function createDataverseChildPaging(entitySetName: string, fetcher: typeof fetch = fetch): (fetchXml: string) => Promise<ChildFetchResult> {
    const entitySet = childLogicalName(entitySetName);
    return async fetchXml => {
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(),20000);
        try {
            const response = await fetcher(`/api/data/v9.2/${entitySet}?fetchXml=${encodeURIComponent(fetchXml)}`, {
                method:'GET',credentials:'same-origin',redirect:'error',signal:controller.signal,
                headers:{Accept:'application/json','OData-MaxVersion':'4.0','OData-Version':'4.0',Prefer:'odata.include-annotations="*"'},
            });
            if (!response.ok) throw new Error(`Unable to load child records (Dataverse HTTP ${response.status}).`);
            const result = await response.json() as {value?: ComponentFramework.WebApi.Entity[]; '@Microsoft.Dynamics.CRM.fetchxmlpagingcookie'?: string; '@Microsoft.Dynamics.CRM.morerecords'?: boolean};
            if (!Array.isArray(result.value)) throw new Error('Dataverse returned an invalid child page.');
            return {entities:result.value,pagingCookie:result['@Microsoft.Dynamics.CRM.fetchxmlpagingcookie'],moreRecords:result['@Microsoft.Dynamics.CRM.morerecords'] === true};
        } catch (error: unknown) {
            if (controller.signal.aborted) throw new Error('The child records request timed out. Refresh the grid to try again.');
            throw error;
        } finally {clearTimeout(timeout);}
    };
}
