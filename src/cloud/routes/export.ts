import { ExportController } from '../controllers';

const definitions: Sensbox.RouteDefinitions = {
  exportMedicalRecords: {
    action: ExportController.exportMedicalRecords,
    secure: true,
  },
};
export default definitions;
