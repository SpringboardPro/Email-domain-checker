// TO: Two external emails of the same domain.
const testCase1 =   [   
                        [
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test2@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]



// TO: Two external emails of different domains.
const testCase2 =   [
                        [
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@yahoo.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]


// TO: One internal and one external email.
const testCase3 =   [
                        [
                            {
                                "emailAddress": "test@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@yahoo.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]


// TO: One internal and two external emails of the same domain.
const testCase4 =   [
                        [
                            {
                                "emailAddress": "test@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test2@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]


// TO: One internal and two external emails of different domains
const testCase5 =   [
                        [
                            {
                                "emailAddress": "test@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@yahoo.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]


// TO: Two internal emails and one external email.
const testCase6 =   [
                        [
                            {
                                "emailAddress": "test@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test2@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]

// TO: Two internal emails and two internal emails of different domains.
const testCase7 =   [
                        [
                            {
                                "emailAddress": "test@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {                            
                                "emailAddress": "test2@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@yahoo.com",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]

// TO: External email.
// CC: External email of same domain.
const testCase8 =   [   
                        [
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                        ],
                        [
                            {
                                "emailAddress": "test2@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]

// TO: External email.
// CC: External email of different domain.
const testCase9 =   [
                        [
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                        ],
                        [
                            {
                                "emailAddress": "test@yahoo.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]

// TO: Internal email.
// CC: External email.
const testCase10 =   [
                        [
                            {
                                "emailAddress": "test@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                        ],
                        [
                            {
                                "emailAddress": "test@yahoo.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]

// TO: Internal email.
// CC: Two external emails of same domains.
const testCase11 =   [
                        [
                            {
                                "emailAddress": "test@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                        ],
                        [
                            {
                                "emailAddress": "test2@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]

// TO: Internal email and external email.
// CC: External email of different domain.
const testCase12 =   [
                        [
                            {
                                "emailAddress": "test@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                        ],
                        [
                            {
                                "emailAddress": "test@yahoo.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]

// TO: Internal email.
// CC: Internal email and external email.
const testCase13 =   [
                        [
                            {
                                "emailAddress": "test@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                        ],
                        [
                            {
                                "emailAddress": "test2@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]

// TO: Internal email.
// CC: Internal email and two external emails of different domain.
const testCase14 =   [
                        [
                            {
                                "emailAddress": "test@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                        ],
                        [
                            {                            
                                "emailAddress": "test2@springboard.pro",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@yahoo.com",
                                "displayName": "test",
                                "recipientType": "other"
                            },
                            {
                                "emailAddress": "test@gmail.com",
                                "displayName": "test",
                                "recipientType": "other"
                            }
                        ]
                    ]



/**
 * Function that a recipient data object and returns just the email addresses as an array.
 * @param {array} result - An array containing the recipient data.
 */
function processEmails (result) {
  let recipientData
  if (result.length > 1) {
    recipientData = result[0].concat(result[1])
  } else {
    recipientData = result[0]
  }
  let emails = []
  for (let i = 0; i < recipientData.length; i++) {
    let Email = recipientData[i].emailAddress
    emails.push(Email)
  }
  return emails
}

test('Returns correct array of emails from recipient object.', () => {
    expect(processEmails(testCase1)).toEqual(["test@gmail.com", "test2@gmail.com"]);
    expect(processEmails(testCase2)).toEqual(["test@gmail.com", "test@yahoo.com"]);
    expect(processEmails(testCase3)).toEqual(["test@springboard.pro", "test@yahoo.com"]);
    expect(processEmails(testCase4)).toEqual(["test@springboard.pro", "test@gmail.com", "test2@gmail.com"]);
    expect(processEmails(testCase5)).toEqual(["test@springboard.pro", "test@gmail.com", "test@yahoo.com"]);
    expect(processEmails(testCase6)).toEqual(["test@springboard.pro", "test2@springboard.pro", "test@gmail.com"]);
    expect(processEmails(testCase7)).toEqual(["test@springboard.pro", "test2@springboard.pro", "test@yahoo.com", "test@gmail.com"]);
    expect(processEmails(testCase8)).toEqual(["test@gmail.com", "test2@gmail.com"]);
    expect(processEmails(testCase9)).toEqual(["test@gmail.com", "test@yahoo.com"]);
    expect(processEmails(testCase10)).toEqual(["test@springboard.pro", "test@yahoo.com"]);
    expect(processEmails(testCase11)).toEqual(["test@springboard.pro", "test@gmail.com", "test2@gmail.com"]);
    expect(processEmails(testCase12)).toEqual(["test@springboard.pro", "test@gmail.com", "test@yahoo.com"]);
    expect(processEmails(testCase13)).toEqual(["test@springboard.pro", "test2@springboard.pro", "test@gmail.com"]);
    expect(processEmails(testCase14)).toEqual(["test@springboard.pro", "test2@springboard.pro", "test@yahoo.com", "test@gmail.com"]);
})

/**
 * Function that returns a boolean value based on if the number of external emails is larger than
 * @param {array} emails - An array containing the emails to be checked.
 */
function checkMultipleExternal (emails) {
  let externalEmails = []
  for (let i = 0; i < emails.length; i++) {
    let domain = emails[i].slice(emails[i].indexOf('@'), emails[i].length)
    if (domain !== '@springboard.pro') {
      externalEmails.push(domain)
    }
  }
  const numberExternalDomains = new Set(externalEmails).size
  if (numberExternalDomains > 1) {
    return true
  } else {
    return false
  }
}

testCase1b = ["test@gmail.com", "test2@gmail.com"]
testCase2b = ["test@gmail.com", "test@yahoo.com"]
testCase3b = ["test@springboard.pro", "test@yahoo.com"]
testCase4b = ["test@springboard.pro", "test@gmail.com", "test2@gmail.com"]
testCase5b = ["test@springboard.pro", "test@gmail.com", "test@yahoo.com"]
testCase6b = ["test@springboard.pro", "test2@springboard.pro", "test@gmail.com"]
testCase7b = ["test@springboard.pro", "test2@springboard.pro", "test@yahoo.com", "test@gmail.com"]

test('Returns correct boolean value for array of email addresses.', () => {
    expect(checkMultipleExternal(testCase1b)).toBe(false);
    expect(checkMultipleExternal(testCase2b)).toBe(true);
    expect(checkMultipleExternal(testCase3b)).toBe(false);
    expect(checkMultipleExternal(testCase4b)).toBe(false);
    expect(checkMultipleExternal(testCase5b)).toBe(true);
    expect(checkMultipleExternal(testCase6b)).toBe(false);
    expect(checkMultipleExternal(testCase7b)).toBe(true); 
})
