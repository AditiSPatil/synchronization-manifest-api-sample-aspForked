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

  const siteId = process.env.SHAREPOINT_SITE_ID!;
  const requestUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
  const drivesResponse = await fetch(requestUrl, {
    method: "GET",
    headers: requestHeaders,
  });

  let driveId: string = "";
  if (drivesResponse.ok) {
    const drivesData = await drivesResponse.json();
    //console.log("Drives fetched successfully:", drivesData);
    const drive = drivesData.value.find(
      (drive: any) =>
        drive.name == "Documents" && drive.driveType === "documentLibrary"
    );

    if (drive) {
      driveId = drive.id;
      console.log(
        "First drive with name containing 'document' and driveType 'documentLibrary':",
        drive.name
      );
    } else {
      console.log("No drive found with the specified criteria.");
    }
  } else {
    console.error("Failed to fetch drives:", drivesResponse.statusText);
  }

  if (!driveId) {
    console.log("Drive ID is empty, returning empty response.");
    return [];
  }
  const items = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`,
    {
      method: "GET",
      headers: requestHeaders,
    }
  );
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
