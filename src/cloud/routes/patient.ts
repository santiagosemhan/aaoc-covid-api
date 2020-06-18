import { PatientController } from '../controllers';

const definitions: Sensbox.RouteDefinitions = {
  fetchLastMedicalRecord: {
    action: PatientController.fetchLastMedicalRecord,
    secure: true,
  },
};
export default definitions;
