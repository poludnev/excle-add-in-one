type TShipper = [string, string, string, string];
type TConsignee = [string, string, string, string];

const shippers: { [key: string]: TShipper } = {
  SAVIMPEX: ["SAVIMPEX DOO", "Novi Sad, Gogoljeva 7, Republic of Serbia", "PIB: 113113438", ""],
};
const consignees: { [key: string]: TConsignee } = {
  URSUS_TRADE: [
    "URSUS TRADE LLC",
    "Russia, Moscow, Hlebniy pereulok 19A, 121069",
    "TIN 7735189429",
    "",
  ],
};

const places_of_delivery: { [key: string]: string } = {
  CROCUS:
    'LLC "URSUS TRADE" / SVH OOO "Crocus Interservice" 143002 Moskovskaya Obl, Odintsovskiy r-n, S. Akulovo, Ul. Novaya, D. 137',
};

const senders_instructions: { [key: string]: string } = {
  AKULOVO:
    'T/P "Akulovskiy" Code 10013010 SVH OOO "Crocus Interservice" 143002 Moskovskaya Obl, Odintsovskiy r-n, S. Akulovo, Ul. Novaya, D. 137      Lic. 10013/200111/10118/11 from 10.10.2024',
};

export { shippers, consignees, places_of_delivery, senders_instructions };
