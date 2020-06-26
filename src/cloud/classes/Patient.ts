import Base from './Base';

class Patient extends Base {
  constructor() {
    super(Patient.name);
  }

  static async beforeSave(request: Parse.Cloud.BeforeSaveRequest) {
    await super.beforeSave(request);
    const { object } = request;
    const acl = object.getACL();
    if (acl) {
      acl.setPublicReadAccess(true);
    }
  }
}

export default Patient;
