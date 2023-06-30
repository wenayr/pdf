

type userDataOther = {
    family: string,
    name: string,
    name2: string,
    email: string,
} | {
    nameOrg: string
}
type userDataMin = {
    tel: number,
    password: string,
    tarifMode: "предоставление услуг"| "безпредоставления услуг"
}
type userData = userDataMin & userDataOther

type tarif = {
    countOrg: number,
    countTS: number,
    dateFrom: Date,
    dateTo: Date,
    stoimist: number, //
    oplata: number //
}
type tPrava = {
    create: boolean,
    refactor: boolean,
    del: boolean
}
type userPrava = {
    createOrg: tPrava,
    voditeli: tPrava,
    users: tPrava,
    sotrudniki: tPrava,
    transport: tPrava,
    dpetvoiDoc: tPrava,
}

interface data {

}


