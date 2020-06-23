/**
 * Search for last record for a given patient for a user. If no record found for patient created by the user,
 * it search for a record created by other users.
 *
 * @param patientId id of patient
 * @param user request user
 */
const fetchLastMedicalRecord = async (
  patientId: string,
  user: Parse.User,
): Promise<Parse.Object | undefined> => {
  if (!patientId) throw new Parse.Error(400, 'patientId is required');

  try {
    const patient = await new Parse.Query('Patient').get(patientId, { useMasterKey: true });
    const lastMedicalRecord = await new Parse.Query('MedicalRecord')
      .include(['morfologia', 'topografia'])
      .equalTo('patient', patient)
      .descending('createdAt')
      .first({ sessionToken: user.getSessionToken() });

    if (lastMedicalRecord) return lastMedicalRecord;
    return new Parse.Query('MedicalRecord')
      .include(['morfologia', 'topografia'])
      .equalTo('patient', patient)
      .descending('createdAt')
      .first({ useMasterKey: true });
  } catch (error) {
    throw new Parse.Error(400, `Cannot fetchLastMedicalRecord. Reason: ${error.message}`);
  }
};

export default {
  fetchLastMedicalRecord,
};
