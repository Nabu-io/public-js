
import { CellHyperlinkValue, Workbook } from 'exceljs'
import { validateEmail, validatePassword } from './validationHelpers';
import { ImportSeedPayload } from './interface';

const GROUP_TYPE = ['SERVICE', 'TEAM', 'OFFICE']
const ROLE = ['USER', 'MANAGER', 'ADMIN', 'SUPERADMIN', 'GUEST']
const POSITION = ['REGISTERED_CUSTOMS_REPRESENTATIVE', 'CUSTOMS_MANAGER', 'CUSTOMS_TEAM_MANAGER', 'INFORMATION_SYSTEMS_MANAGER', 'NABU_ADMINISTRATOR']
const LANGUAGE = ['fr', 'en']

const STREET_TYPE: string[] = [
    "Street", "St", "Road", "Rd", "Avenue", "Ave", "Boulevard", "Blvd",
    "Drive", "Dr", "Lane", "Ln", "Court", "Ct", "Plaza", "Plz", "Square", "Sq",
    "Terrace", "Ter", "Place", "Pl", "Trail", "Trl", "Way", "Wy", "Loop", "Lp",
    "Alley", "Aly", "Parkway", "Pkwy", "Esplanade", "Expressway", "Expwy",
    "Freeway", "Fwy", "Highway", "Hwy", "Circle", "Cir", "Close", "Cl",
    "Crescent", "Cres", "Drive", "Dr", "Gardens", "Gdns", "Gate", "Gt",
    "Green", "Grn", "Grove", "Grv", "Hill", "Hl", "Island", "Isld",
    "Junction", "Jct", "Key", "Ky", "Landing", "Lndg", "Meadow", "Mdw",
    "Mews", "Pass", "Path", "Pike", "Row", "Rue", "Run", "Walk", "Quay",
    "Crossing", "Xing", "Circle", "Crcle", "Corridor", "Cory", "Arcade", "Arc",
    "Bay", "By", "Beach", "Bch", "Bend", "Bnd", "Cape", "Cpe", "Cliff", "Clf",
    "Common", "Cmn", "Corner", "Cor", "Camp", "Cp", "Curve", "Cv", "Cove", "Cv",
    "Dale", "Dl", "Dam", "Dm", "Divide", "Dv", "Dock", "Dk", "Estate", "Est",
    "Flats", "Flts", "Ford", "Frd", "Forest", "Frst", "Fork", "Frk", "Fort", "Ft",
    "Glen", "Gln", "Harbor", "Hbr", "Haven", "Hvn", "Heights", "Hts", "Highway", "Hwy",
    "Hollow", "Holw", "Inlet", "Inlt", "Island", "Isl", "Junction", "Jct", "Knoll", "Knl",
    "Lake", "Lk", "Land", "Lnd", "Ledge", "Ldg", "Manor", "Mnr", "Mill", "Ml", "Mission", "Msn",
    "Mount", "Mt", "Mountain", "Mtn", "Orchard", "Orch", "Parade", "Pde", "Peak", "Pk",
    "Pines", "Pnes", "Point", "Pt", "Port", "Prt", "Ridge", "Rdg", "River", "Rvr", "Shore", "Shr",
    "Spring", "Spg", "Square", "Sq", "Station", "Stn", "Stravenue", "Stra", "Stream", "Stm",
    "Street", "St", "Summit", "Smt", "Terrace", "Ter", "Turnpike", "Tpke", "Valley", "Vly",
    "Village", "Vlg", "Vista", "Vis", "Walk", "Wlk", "Wall", "Wl", "Way", "Wy", "Wharf", "Whrf",
    "Wood", "Wd", "Woods", "Wds"
];

function parseOrganizationWorksheet(workbook: Workbook): any[] {
    const organizationWorksheet = workbook.getWorksheet('Organizations')
    const organizations: any[] = []
    organizationWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const name = row.getCell(1).value?.toString().trim()
            try {
                if (!name) {
                    throw new Error("Invalid organizationName")
                }
            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Organization)`)
            }

            organizations.push({
                name
            })
        }
    });
    return organizations
}

function parseGroupWorksheet(workbook: Workbook): any[] {
    const groupsWorksheet = workbook.getWorksheet('Groups')
    const groups: any[] = []
    groupsWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const name = row.getCell(1).value?.toString().trim()
            const type = row.getCell(2).value?.toString().trim()
            const parent = row.getCell(3).value?.toString().trim()
            const organizationName = row.getCell(4).value?.toString().trim()
            const valueCompanyRegistrationNumber = row.getCell(5).value?.toString().trim()
            const country = row.getCell(6).value?.toString().trim()
            const addressGivenName = row.getCell(7).value?.toString().trim()
            const addressSurname = row.getCell(8).value?.toString().trim()
            const addressCompanyName = row.getCell(9).value?.toString().trim()
            const addressStreetNo = row.getCell(10).value?.toString().trim()
            const addressStreetType = row.getCell(11).value?.toString().trim()
            const addressStreetName = row.getCell(12).value?.toString().trim()
            const addressFloor = row.getCell(13).value?.toString().trim()
            const addressTown = row.getCell(14).value?.toString().trim()
            const addressRegion = row.getCell(15).value?.toString().trim()
            const addressPostcode = row.getCell(16).value?.toString().trim()
            const addressCountry = row.getCell(17).value?.toString().trim()
            const addressTag = row.getCell(18).value?.toString().trim()

            try {
                if (!name) {
                    throw new Error("Invalid name")
                }

                if (!GROUP_TYPE.includes(type as string)) {
                    throw new Error("Invalid groupType")
                }

                if (!organizationName) {
                    throw new Error("Invalid organizationName")
                }

                if (!valueCompanyRegistrationNumber) {
                    throw new Error("Invalid companyRegistrationNumber")
                }

                if (!country) {
                    throw new Error("Invalid country")
                }

                if (!addressGivenName) {
                    throw new Error("Invalid addressGivenName")
                }

                if (!addressSurname) {
                    throw new Error("Invalid addressSurname")
                }

                if (!addressCompanyName) {
                    throw new Error("Invalid addressCompanyName")
                }

                if (!addressStreetNo) {
                    throw new Error("Invalid addressStreetNo")
                }

                if (!STREET_TYPE.includes(addressStreetType as string)) {
                    throw new Error("Invalid addressStreetType")
                }

                if (!addressStreetName) {
                    throw new Error("Invalid addressStreetName")
                }

                if (!addressFloor) {
                    throw new Error("Invalid addressFloor")
                }

                if (!addressTown) {
                    throw new Error("Invalid addressTown")
                }

                if (!addressRegion) {
                    throw new Error("Invalid addressRegion")
                }

                if (!addressPostcode) {
                    throw new Error("Invalid addressPostcode")
                }


                if (!addressCountry) {
                    throw new Error("Invalid addressCountry")
                }

                if (!addressTag) {
                    throw new Error("Invalid addressTag")
                }

            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Groups)`)
            }

            groups.push({
                name,
                type,
                parent,
                organizationName,
                companyRegistrationNumber: {
                    value: valueCompanyRegistrationNumber,
                    country: country
                },
                address: {
                    givenName: addressGivenName,
                    surname: addressSurname,
                    companyName: addressCompanyName,
                    streetNo: addressStreetNo,
                    streetType: addressStreetType,
                    streetName: addressStreetName,
                    floor: addressFloor,
                    town: addressTown,
                    region: addressRegion,
                    postCode: addressPostcode,
                    country: addressCountry,
                    tag: addressTag
                }
            })
        }
    });
    return groups
}

function parseUsersWorksheet(workbook: Workbook, groupsName: string[]): any[] {
    const usersWorksheet = workbook.getWorksheet('Users')
    const users: any[] = []
    usersWorksheet?.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 1) {
            const email = (row.getCell(1).value as CellHyperlinkValue).text.trim()
            const role = row.getCell(2).value?.toString().trim()
            const active = row.getCell(3).value?.toString().trim()
            const groupName = row.getCell(4).value?.toString().trim()
            const firstName = row.getCell(5).value?.toString().trim()
            const lastName = row.getCell(6).value?.toString().trim()
            const position = row.getCell(7).value?.toString().trim()
            const language = row.getCell(8).value?.toString().trim()
            const password = row.getCell(9).value?.toString().trim()
            const sendEmail = row.getCell(10).value?.toString().trim()

            try {
                if (!validateEmail(email)) {
                    throw new Error("Invalid email")
                }

                if (!ROLE.includes(role as string)) {
                    throw new Error("Invalid role")
                }

                if (!['oui', 'non'].includes(active as string)) {
                    throw new Error("Invalid value for active")
                }

                if (!groupName) {
                    throw new Error("Invalid groupName")
                }

                if (!groupsName.includes(groupName)) {
                    throw new Error("Group name is not defined in the previous worksheet")
                }

                if (!firstName) {
                    throw new Error("Invalid first name")
                }

                if (!lastName) {
                    throw new Error("Invalid last name")
                }

                if (!POSITION.includes(position as string)) {
                    throw new Error("Invalid position")
                }

                if (!LANGUAGE.includes(language as string)) {
                    throw new Error("Invalid language")
                }

                if (!validatePassword(password as string)) {
                    throw new Error("Invalid password format")
                }

                if (!['oui', 'non'].includes(sendEmail as string)) {
                    throw new Error("Invalid value for sendEmail expected 'oui' or 'non'")
                }

            } catch (err) {
                throw new Error(`${err} at rows number: ${rowNumber} (Worksheeet: Users)`)
            }

            users.push({
                email,
                role,
                active: active === "oui" ? true : false,
                groupName,
                firstName,
                lastName,
                position,
                language,
                password,
                sendEmail: sendEmail === "oui" ? true : false
            })
        }
    })
    return users
}

export default function processWorkbook(workbook: any) {
    /**
     * Parse the excel file
     */
    const jsonData: ImportSeedPayload = {
        organizations: [],
        groups: [],
    };



    jsonData.organizations = parseOrganizationWorksheet(workbook);
    const groups = parseGroupWorksheet(workbook)



    const groupsName = groups.map((group) => group.name as string)
    const users = parseUsersWorksheet(workbook, groupsName);

    jsonData.groups = groups;

    jsonData.groups.forEach(group => {
        const usersInGroup = users.filter(user => user.groupName === group.name);
        usersInGroup.forEach(user => delete user.groupName);
        group.users = usersInGroup;
    });

    return jsonData;
}
