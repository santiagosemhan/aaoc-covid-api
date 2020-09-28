import { RolesType } from '../../types/constants';
import SecurityService from './SecurityService';

const fetchMedicalRecords = async (user: Parse.User) => {
  const LIMIT_RECORDS = 10000;

  const isAdmin = await SecurityService.hasUserRole(user, RolesType.ADMINISTRATOR);
  const pipeline = [];
  if (!isAdmin) {
    pipeline.push({ match: { _p_createdBy: `_User$${user.id}` } });
  }

  pipeline.push(
    ...[
      { sort: { _created_at: -1 } },
      {
        group: {
          objectId: '$patient',
          id: { $first: '$_id' },
        },
      },
      { limit: LIMIT_RECORDS },
    ],
  );

  // @ts-ignore
  const results = await new Parse.Query('MedicalRecord').aggregate(pipeline, {
    useMasterKey: true,
  });
  const medicalRecordsIds = results.map((mr: { id: string }) => mr.id);

  const medicalRecords = await new Parse.Query('MedicalRecord')
    .include([
      'patient.paisNacimiento',
      'morfologia',
      'topografia',
      'createdBy.account.organization',
    ])
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
