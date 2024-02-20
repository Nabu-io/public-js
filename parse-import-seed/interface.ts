

export interface User {
    email: string;
    role: string;
    active: boolean;
    firstName: string;
    lastName: string;
    position: string;
    language: string;
    password: string;
    sendEmail: boolean;
}

export interface CompanyRegistrationNumber {
    value: string;
    country: string;
}

export interface Address {
    givenName: string;
    surname: string;
    companyName: string;
    streetNo: string;
    streetName: string;
    streeType: string;
    floor: string;
    town: string;
    region: string;
    postCode: string;
    country: string;
    tag: string;
}

interface Organization {
    name: string
}

interface Group {
    name: string;
    type: string;
    parent: string | null
    organizationName: string;
    users: User[],
    companyRegistrationNumber: CompanyRegistrationNumber,
    address: Address,
}

export interface ImportSeedPayload {
    organizations: Organization[];
    groups: Group[];
}
