export interface RefereeDetails {
  name: string;
  title: string;
  company: string;
  stature: string;
  address: string;
  email: string;
  phone: string;
  linkedin: string;
}

export interface ApplicantDetails {
  name: string;
  role: string;
  company: string;
  visaRoute: "Exceptional Talent" | "Exceptional Promise";
  selectedCriteria: string[]; // ['MC', 'OC1', 'OC2', 'OC3', 'OC4']
  relationship: string;
  anecdotes: string;
  customDate: string;
}

export interface PresetCase {
  id: string;
  name: string;
  description: string;
  referee: RefereeDetails;
  applicant: ApplicantDetails;
}
