const fetchMedicalRecords = async (user: Parse.User) => {
  const LIMIT_RECORDS = 10000;
  const pipeline = [
    { sort: { createdAt: -1 } },
    {
      group: {
        objectId: '$patient',
        id: { $last: '$_id' },
      },
    },
    { limit: LIMIT_RECORDS },
  ];

  // @ts-ignore
  const results = await new Parse.Query('MedicalRecord').aggregate(pipeline, {
    useMasterKey: true,
  });
  const medicalRecordsIds = results.map((mr: { id: string }) => mr.id);

  const medicalRecords = await new Parse.Query('MedicalRecord')
    .include(['patient', 'morfologia', 'topografia'])
    .containedIn('objectId', medicalRecordsIds)
    .limit(LIMIT_RECORDS)
    .find({
      sessionToken: user.getSessionToken(),
    });

  return medicalRecords;
};

export default {
  fetchMedicalRecords,
};
