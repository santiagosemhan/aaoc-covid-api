import ExportService from '../services/ExportService';

const exportMedicalRecords = (
  request: Sensbox.SecureFunctionRequest,
): Promise<Parse.Object | undefined> => {
  const { user } = request;
  return ExportService.exportMedicalRecords(user);
};

export default {
  exportMedicalRecords,
};
