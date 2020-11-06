//  Copyright (c) Microsoft. All rights reserved.
//  Licensed under the MIT license.

import React from 'react';
import {
  ActivityIndicator,
  Alert,
  FlatList,
  Modal,
  StyleSheet,
  Text,
  View,
  ScrollView
} from 'react-native';
import { createStackNavigator } from '@react-navigation/stack';
import { Table, Row, Rows } from 'react-native-table-component';
import moment from 'moment';

import { DrawerToggle, headerOptions } from '../menus/HeaderComponents';
import { GraphManager } from '../graph/GraphManager';

const Stack = createStackNavigator();
const initialState: WorkbookScreenState = { loadingData: true, data: {} };
const WorkbookState = React.createContext(initialState);

type WorkbookScreenState = {
  loadingData: boolean;
  data: any;
}

// Temporary JSON view
const WorkbookComponent = () => {
  const workbookState = React.useContext(WorkbookState);

  return (
    <ScrollView style={styles.container}>
      <Modal visible={workbookState.loadingData}>
        <View style={styles.loading}>
          <ActivityIndicator animating={workbookState.loadingData} size='large' />
        </View>
      </Modal>
      <View style={styles.section}>
        <Text style={styles.sectionTitle}>Worksheets</Text>
        {workbookState.data.worksheets?.map((item:any) => 
          <View style={styles.item} key={item.name}>
            <Text style={styles.name}>{item.name}</Text>
          </View>
        )}
      </View>
      <View style={styles.section}>
        <Text style={styles.sectionTitle}>Tables</Text>
        {workbookState.data.tables?.map((item:any) => 
          <View style={styles.item} key={item.name}>
            <Text style={styles.name}>{item.name}</Text>
          </View>
        )}
      </View>
      <View style={styles.section}>
        <Text style={styles.sectionTitle}>Calculate Loan Payment</Text>
        <Text style={styles.item}>Rate: {workbookState.data.loan?.rate}</Text>
        <Text style={styles.item}>Number of periods (nper): {workbookState.data.loan?.nper}</Text>
        <Text style={styles.item}>Present value (pv): {workbookState.data.loan?.pv}</Text>
        <Text style={styles.item}>Loan Payment (pmt): {workbookState.data.loan?.pmt}</Text>
      </View>
      <View style={styles.section}>
        <Text style={styles.sectionTitle}>Get Range</Text>
        <Text style={styles.item}>Address: {workbookState.data.range?.address}</Text>
        <Table borderStyle={{borderWidth: 1, borderColor: '#c8e1ff'}}>
          <Rows data={workbookState.data.range?.values} textStyle={styles.cell}/>
        </Table>
      </View>
      <View style={styles.section}>
        <Text style={styles.sectionTitle}>Calculate Profit</Text>
        <Text style={styles.item}>Address: {workbookState.data.profitRange?.address}</Text>
        <Table borderStyle={{borderWidth: 1, borderColor: '#c8e1ff'}}>
          <Rows data={workbookState.data.profitRange?.values} textStyle={styles.cell}/>
        </Table>
      </View>
    </ScrollView>
  );
}

export default class WorkbookScreen extends React.Component {

  state: WorkbookScreenState = initialState;

  async componentDidMount() {
    try {
      // Copy the Excel workbook GraphTutorial.xlsx from the Excel folder in this project to OneDrive
      // Replace the workbook id here with the id of the GraphTutorial.xslx on OneDrive (use Microsoft Graph Explorer to find it)
      const workbookId = '01JLZJVCRQBG5OWCHLLNCIVPTUVRPQ466M';

      const workbookData: any = {};

      // Get worksheets
      workbookData.worksheets = (await GraphManager.getWorksheets(workbookId)).value;

      // Get tables
      workbookData.tables = (await GraphManager.getTables(workbookId)).value;

      // Calculate loan payment
      workbookData.loan = {
        "rate": 0.035,
        "nper": 20,
        "pv": -2000
      };
      workbookData.loan.pmt = (await GraphManager.calculateLoanPayment(
        workbookId, 
        workbookData.loan.rate, 
        workbookData.loan.nper, 
        workbookData.loan.pv
      )).value;

      // Get range
      workbookData.range = await GraphManager.getRange(workbookId, 'Product Catalog', 'B4:C7');

      // Calculate profit
      await GraphManager.setRange(workbookId, 'Profit Calculator', 'C4:C5', 
        { values: [
          [ 6000 ],
          [ 4500 ]
        ]});
      workbookData.profitRange = await GraphManager.getRange(workbookId, 'Profit Calculator', 'B4:C7');

      console.log(workbookData);

      this.setState({
        loadingData: false,
        data: workbookData
      });
    } catch(error) {
      Alert.alert(
        'Error getting workbook data',
        JSON.stringify(error),
        [
          {
            text: 'OK'
          }
        ],
        { cancelable: false }
      );
    }
  }

  render() {
    return (
      <WorkbookState.Provider value={this.state}>
        <Stack.Navigator screenOptions={ headerOptions }>
          <Stack.Screen name='Workbook'
            component={ WorkbookComponent }
            options={{
              title: 'Workbook',
              headerLeft: () => <DrawerToggle/>
            }} />
        </Stack.Navigator>
      </WorkbookState.Provider>
    );
  }
}

const styles = StyleSheet.create({
  container: {
    margin: 10
  },
  loading: {
    flex: 1,
    justifyContent: 'center',
    alignItems: 'center'
  },
  section: {
    flex: 1,
    paddingTop: 10,
    paddingBottom: 10
  },
  sectionTitle: {
    fontWeight: '700',
    fontSize: 18
  },
  item: {
    paddingTop: 5,
    paddingBottom: 5
  },
  name: {
    fontWeight: '200',
  },
  cell: {
    margin: 2
  }
});
