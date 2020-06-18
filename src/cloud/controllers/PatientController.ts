import PatientService from '../services/PatientService';

const fetchLastMedicalRecord = async (
  request: Sensbox.SecureFunctionRequest,
): Promise<Parse.Object | undefined> => {
  const { params, user } = request;
  const { patientId } = params;
  return PatientService.fetchLastMedicalRecord(patientId, user);
};

export default {
  fetchLastMedicalRecord,
};
