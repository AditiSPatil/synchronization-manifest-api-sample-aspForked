/*---------------------------------------------------------------------------------------------
 * Copyright (c) Bentley Systems, Incorporated. All rights reserved.
 * See LICENSE.md in the project root for license terms and full copyright notice.
 *--------------------------------------------------------------------------------------------*/

import { ContainerClient } from "@azure/storage-blob";
import { getAccessUrl, getMSAccessToken } from "./externalAccess";
import fetch from "node-fetch";

/*
 * Get items from Azure Storage
 */
export const getAzureStorageItems = async () => {
  const accessUrl = getAccessUrl("");
  const containerClient: ContainerClient = new ContainerClient(accessUrl);
  const azureFiles = [];

  for await (const blob of containerClient.listBlobsFlat()) {
    azureFiles.push({
      id:
        containerClient.accountName +
        "/" +
        containerClient.containerName +
        "/" +
        blob.name,
      isFolder: false,
      name: blob.name,
      modifiedDate: blob.properties.lastModified,
      downloadUrl: getAccessUrl(blob.name),
    });
  }

  return azureFiles;
};

/*
 * Get items from SharePoint Storage
 */
export const getSharePointStorageItems = async () => {
  const accessObject = await getMSAccessToken();

  const requestHeaders = {
    Authorization: "Bearer " + accessObject!.accessToken,
    "Content-Type": "application/json",
  };

  console.log(accessObject); //asp
  const requestUrl =
    "https://graph.microsoft.com/v1.0/sites/bentley.sharepoint.com,e809a1f6-23de-4991-abc0-a2db14119a1f,dfe70037-c4e8-4df5-9d38-ae22e07fe948/drives"; //"https://graph.microsoft.com/v1.0/me/drive";
  const drivesResponse = await fetch(requestUrl, {
    method: "GET",
    headers: requestHeaders,
  });

  const drives = await drivesResponse.json();
  console.log(drives);
  const driveId = drives.value[0].id as string;
  console.log("asp", drives.value[0].id);
  console.log(driveId);
  const items = await fetch(
    "	https://graph.microsoft.com/v1.0/drives/b!9qEJ6N4jkUmrwKLbFBGaHzcA59_oxPVNnTiuIuB_6UgG9XnrDIwiRoYhrjdEWGDJ/items/01IHA6PSCGXUFOWYD5IBFZFA33XAYESA5U/children", //basic working`https://graph.microsoft.com/v1.0/sites/e809a1f6-23de-4991-abc0-a2db14119a1f/drive/root/children`, //original `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`,
    {
      method: "GET",
      headers: requestHeaders,
    }
  );
  console.log(items);
  const result = await items.json();

  return result.value.map((item: any) => {
    return {
      id: item.id,
      name: item.name,
      isFolder: item.folder ? true : false,
      modifiedDate: item.lastModifiedDateTime,
      downloadUrl: item["@microsoft.graph.downloadUrl"],
    };
  });
};
