function GoogleContacts() {}

GoogleContacts.prototype.getContacts = function () {
    var contacts = ContactsApp.getContacts();
    var contacts_formatted = contacts.map(function (google_contact) {
        return {
            fullname: google_contact.getFullName(),
            phones: JSON.stringify(google_contact.getPhones().map(function (phone) {
                return {
                    label: phone.getLabel(),
                    number: phone.getPhoneNumber()
                }
            })),
            email: google_contact.getPrimaryEmail(),
            emails: JSON.stringify(google_contact.getEmails().map(function(email) {
                return {
                    label: email.getLabel(),
                    email: email.getAddress()
                }
            })),
            dates: JSON.stringify(google_contact.getDates().map(function(date) {
                return {
                    label: date.getLabel(),
                    email: date.getMonth() + "/" + date.getDay() + "/" + date.getYear()
                }
            })) 
        }
    });
    return contacts_formatted;
}