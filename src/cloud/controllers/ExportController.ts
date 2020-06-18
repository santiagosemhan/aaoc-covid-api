import ExportService from '../services/ExportService';

const exportMedicalRecords = async (
  request: Sensbox.SecureFunctionRequest,
): Promise<Parse.Object> => {
  const { user } = request;
  const exportedMedicalRecords = await ExportService.exportMedicalRecords(user);
  return exportedMedicalRecords;
};

export default {
  exportMedicalRecords,
};
