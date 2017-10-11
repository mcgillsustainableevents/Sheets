const google = require('googleapis');
const sheets = google.sheets('v4');
const OAuth2 = google.auth.OAuth2;
const oauth2Client = new OAuth2(
  '377631274120-2ap2u4ock2t4ig92kpr04lseatlvnh8k.apps.googleusercontent.com',
  'BwysxhUL8CGl2nFJHzQlIj37'
);
oauth2Client.setCredentials({ refresh_token: '1/xquD_gVlBnx7CROe6fuhFjTdQR6QATxDShfMqL2GVOE' });

exports.handler = (req, res) => {
  if (req.method === 'OPTIONS') {
    res
      .set('Access-Control-Allow-Origin', '*')
      .set('Access-Control-Allow-Methods', 'POST')
      .set('Access-Control-Allow-Headers', 'Content-Type')
      .sendStatus(200);
    return;
  }

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
  } = req.body;

  const request = {
    auth: oauth2Client,
    spreadsheetId: '1ou9XWDFg7bgBq3kt_ZOZfSVHgOMsHQc1sd5HulzXSlQ',
    resource: {
      requests: [
        {
          insertDimension: {
            range: {
              sheetId: 0,
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
                      stringValue: amount
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
                  },
                  {
                    userEnteredValue: {
                      stringValue: checked
                    }
                  }
                ]
              }
            ],
            fields: '*',
            start: {
              sheetId: 0,
              rowIndex: 1,
              columnIndex: 2
            }
          }
        }
      ]
    }
  };

  oauth2Client.refreshAccessToken(() => {
    sheets.spreadsheets.batchUpdate(request, () =>
      res.set('Access-Control-Allow-Origin', '*').sendStatus(200)
    );
  });
};
