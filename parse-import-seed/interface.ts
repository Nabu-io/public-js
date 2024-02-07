export interface Group {
    groupName: string;
    groupType: string;
    parent: string | null;
}

export interface User {
    email: string;
    role: string;
    active: boolean;
    groupName: string;
    firstName: string;
    lastName: string;
    position: string;
    language: string;
    password: string;
}

export interface CompanyRegistrationNumber {
    value: string;
    country: string;
    groupName: string;
}

export interface Address {
    givenName: string;
    surname: string;
    companyName: string;
    streetNumber: string;
    streetName: string;
    floor: string;
    town: string;
    region: string;
    postCode: string;
    country: string;
    tag: string;
    groupName: string;
}

export interface Schema {
    name: string;
    content: any; // Can be more specific based on your JSON structure
}

export interface ExportDataTemplate {
    name: string;
    schemaName: string;
    isNew: boolean;
}

export interface Transport {
    type: string;
    name: string;
    address: string;
    port: string;
    username: string;
    password: string;
}

export interface Configuration {
    name: string;
    exportDataTemplateName: string;
    transportName: string;
    groupName: string;
    preset: string;
}

export interface Setting {
    key: string;
    value: string;
    groupName: string;
}

export interface ImportSeedPayload {
    organizations: Group[];
    groups: Group[];
    users: User[];
    companyRegistrationNumbers: CompanyRegistrationNumber[];
    exportDataTemplates: ExportDataTemplate[];
    addresses: Address[];
    schemas: Schema[];
    transports: Transport[];
    configurations: Configuration[];
    settings: Setting[];
}
