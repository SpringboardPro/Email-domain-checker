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
                        ],
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
                        ],
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
                        ],
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
                        ],
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
                        ],
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
                        ],
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
                        ],
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
                                "emailAddress": "test2@yahoo.com",
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
                                "emailAddress": "test@springboard.pro",
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

const processEmails = require(./processEmails)


test('Properly returns list of emails from recipient object.', () => {
    expect(processEmails(testCase1[0].concat(testCase1[1])).toBe(["test@gmail.com", "test2@gmail.com"]);
    expect(processEmails(testCase2[0].concat(testCase2[1])).toBe(["test@gmail.com", "test@yahoo.com"]);
    expect(processEmails(testCase3[0].concat(testCase3[1])).toBe(["test@springboard.pro", "test@yahoo.com"]);
    expect(processEmails(testCase4[0].concat(testCase4[1])).toBe(["test@springboard.pro", "test@gmail.com", "test2@gmail.com"]);
    expect(processEmails(testCase5[0].concat(testCase5[1])).toBe(["test@springboard.pro", "test@gmail.com", "test@yahoo.com"]);
    expect(processEmails(testCase6[0].concat(testCase6[1])).toBe(["test@springboard.pro", "test2@springboard.com", "test@gmail.com"]);
    expect(processEmails(testCase7[0].concat(testCase7[1])).toBe(["test@springboard.pro", "test2@springboard.pro", "test@yahoo.com", "test@gmail.com"]);
    expect(processEmails(testCase8[0].concat(testCase8[1])).toBe(["test@gmail.com", "test2@gmail.com"]);
    expect(processEmails(testCase9[0].concat(testCase9[1])).toBe(["test@gmail.com", "test@yahoo.com"]);
    expect(processEmails(testCase10[0].concat(testCase10[1])).toBe(["test@springboard.pro", "test@yahoo.com"]);
    expect(processEmails(testCase11[0].concat(testCase11[1])).toBe(["test@springboard.pro", "test@gmail.com", "test2@gmail.com"]);
    expect(processEmails(testCase12[0].concat(testCase12[1])).toBe(["test@springboard.pro", "test@gmail.com", "test@yahoo.com"]);
    expect(processEmails(testCase13[0].concat(testCase13[1])).toBe(["test@springboard.pro", "test2@springboard.com", "test@gmail.com"]);
    expect(processEmails(testCase14[0].concat(testCase14[1])).toBe(["test@springboard.pro", "test2@springboard.pro", "test@yahoo.com", "test@gmail.com"]);
})