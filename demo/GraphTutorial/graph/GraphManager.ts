//  Copyright (c) Microsoft. All rights reserved.
//  Licensed under the MIT license.

import { Client } from '@microsoft/microsoft-graph-client';

import { GraphAuthProvider } from './GraphAuthProvider';

// Set the authProvider to an instance
// of GraphAuthProvider
const clientOptions = {
  authProvider: new GraphAuthProvider()
};

// Initialize the client
const graphClient = Client.initWithMiddleware(clientOptions);

export class GraphManager {
  static getUserAsync = async() => {
    // GET /me
    return graphClient.api('/me').get();
  }

  // <GetEventsSnippet>
  static getEvents = async() => {
    // GET /me/events
    return graphClient.api('/me/events')
      // $select='subject,organizer,start,end'
      // Only return these fields in results
      .select('subject,organizer,start,end')
      // $orderby=createdDateTime DESC
      // Sort results by when they were created, newest first
      .orderby('createdDateTime DESC')
      .get();
  }
  // </GetEventsSnippet>

  // Get worksheets in workbook
  static getWorksheets = async(workbookId: string) => {
    // GET me/drive/items/{drive-item-id}/workbook/worksheets
    return graphClient.api(`me/drive/items/${workbookId}/workbook/worksheets`)
      .get();
  }

  // Get tables in workbook
  static getTables = async(workbookId: string) => {
    // GET me/drive/items/{drive-item-id}/workbook/tables
    return graphClient.api(`me/drive/items/${workbookId}/workbook/tables`)
      .get();
  }

  // Get range
  static getRange = async(workbookId: string, worksheetId: string, address: string) => {
    // GET me/drive/items/{drive-item-id}/workbook/worksheets/{worksheetId}/range(address='{address}')
    return graphClient.api(`me/drive/items/${workbookId}/workbook/worksheets/${worksheetId}/range(address='${address}')`)
      .get();
  }
  
  // Set range
  static setRange = async(workbookId: string, worksheetId: string, address: string, values: any) => {
    // PATCH me/drive/items/{drive-item-id}/workbook/worksheets/{worksheetId}/range(address='{address}')
    return graphClient.api(`me/drive/items/${workbookId}/workbook/worksheets/${worksheetId}/range(address='${address}')`)
      .patch(values);
  }

  // Calculate loan payment
  static calculateLoanPayment = async(workbookId: string, rate: number, nper: number, pv: number) => {
    // POST /me/drive/items/{drive-item-id}/workbook/functions/pmt
    return graphClient.api(`me/drive/items/${workbookId}/workbook/functions/pmt`)
      .post({
        "rate": rate,
        "nper": nper,
        "pv": pv
    });
  }

}
