exports.handler = (req, res) => {
  console.log(req);
  const {
    amount,
    charging,
    checked,
    date,
    discussions,
    email,
    eventName,
    food,
    location,
    materials,
    name,
    score,
    sponsored,
    supplies
  } = JSON.parse(req.body);

  var google = require('googleapis');
  var OAuth2 = google.auth.OAuth2;
  var oauth2Client = new OAuth2(
    '217090774948-43ecud917atupn5oah4n7d2n0cbahdjk.apps.googleusercontent.com',
    'q7zs-o9Gcl4BeIDQMExM8OrV'
  );

  oauth2Client.setCredentials({
    refresh_token: '1/BAu1PVACzaZMIR_m9xTXOL-IwxBJEwU1Vsfw2TwcTIU'
  });

  oauth2Client.refreshAccessToken((err, tokens) => {
    var sheets = google.sheets('v4');
    var request = {
      auth: oauth2Client,
      spreadsheetId: '1PjJWQAnZhq4oy1OF2yKclt5LWy79eX5DZOm_dzdW36Y',
      resource: {
        requests: [
          {
            insertDimension: {
              range: {
                sheetId: 311459653,
                dimension: 'ROWS',
                startIndex: 1,
                endIndex: 2
              },
              inheritFromBefore: false
            }
          },
          {
            updateCells: {
              rows: [
                {
                  values: [
                    {
                      userEnteredValue: {
                        stringValue: name
                      }
                    },
                    {
                      userEnteredValue: {
                        stringValue: email
                      }
                    },
                    {
                      userEnteredValue: {
                        stringValue: eventName
                      }
                    },
                    {
                      userEnteredValue: {
                        stringValue: date
                      }
                    },
                    {
                      userEnteredValue: {
                        stringValue: location
                      }
                    },
                    {
                      userEnteredValue: {
                        numberValue: amount
                      }
                    },
                    {
                      userEnteredValue: {
                        boolValue: charging
                      }
                    },
                    {
                      userEnteredValue: {
                        boolValue: materials
                      }
                    },
                    {
                      userEnteredValue: {
                        boolValue: discussions
                      }
                    },
                    {
                      userEnteredValue: {
                        boolValue: food
                      }
                    },
                    {
                      userEnteredValue: {
                        boolValue: supplies
                      }
                    },
                    {
                      userEnteredValue: {
                        boolValue: sponsored
                      }
                    },
                    {
                      userEnteredValue: {
                        stringValue: score
                      }
                    }
                  ]
                }
              ],
              fields: '*',
              start: {
                sheetId: 311459653,
                rowIndex: 1,
                columnIndex: 2
              }
            }
          }
        ]
      }
    };
    sheets.spreadsheets.batchUpdate(request, (err, response) => {
      res.status(200).send(
        JSON.stringify({
          body: 'Success!'
        })
      );
    });
  });
};
