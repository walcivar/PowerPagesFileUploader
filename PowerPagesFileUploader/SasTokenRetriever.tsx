const getSasTokenFromDataverse = async (
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    context: ComponentFramework.Context<any>,
    // eslint-enable-next-line @typescript-eslint/no-explicit-any
    environmentVariableName: string
  ): Promise<string> => {
    const fetchXml = `
      <fetch version="1.0" output-format="xml-platform" mapping="logical" distinct="false">
        <entity name="environmentvariablevalue">
          <attribute name="environmentvariablevalueid" />
          <attribute name="value" />
          <attribute name="createdon" />
          <order attribute="createdon" descending="true" />
          <link-entity name="environmentvariabledefinition" from="environmentvariabledefinitionid" to="environmentvariabledefinitionid" link-type="inner" alias="ac">
            <filter type="and">
              <condition attribute="schemaname" operator="eq" value="${environmentVariableName}" />
            </filter>
          </link-entity>
        </entity>
      </fetch>`;
  
    return new Promise((resolve, reject) => {
      context.webAPI
        .retrieveMultipleRecords("environmentvariablevalue", `?fetchXml=${encodeURIComponent(fetchXml)}`)
        .then(
          (response) => {
            if (response.entities.length > 0) {
              resolve(response.entities[0].value);
            } else {
              reject("Environment variable value not found.");
            }
          },
          (error) => {
            reject(error.message);
          }
        );
    });
  };
  
  const getSasTokenFromPowerPages = async (
    powerPagesUrl: string,
    environmentVariableName: string
  ): Promise<string> => {
    const response = await fetch(`${powerPagesUrl}/_api/environmentvariablevalue?$expand=environmentvariabledefinition($filter=schemaname eq '${environmentVariableName}')`, {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
      },
    });
  
    if (!response.ok) {
      throw new Error(`Power Pages API call failed with status: ${response.status}`);
    }
  
    const data = await response.json();
    if (data.value.length > 0) {
      return data.value[0].value;
    } else {
      throw new Error("Environment variable value not found in Power Pages.");
    }
  };
  
  export const getSasToken = async (
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    context: ComponentFramework.Context<any>,
    // eslint-enable-next-line @typescript-eslint/no-explicit-any
    environmentVariableName: string,
    powerPagesUrl: string
  ): Promise<string> => {
    try {
      return await getSasTokenFromDataverse(context, environmentVariableName);
    } catch (error) {
      console.warn(`Dataverse API failed: ${error}`);
      console.info("Attempting to retrieve SAS token using Power Pages API...");
      return getSasTokenFromPowerPages(powerPagesUrl, environmentVariableName);
    }
  };