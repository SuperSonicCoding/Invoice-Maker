// Handles initial startup
console.log('window', window.location.href.split('/').pop());
// Gets the last thing in the file location and uses it to see if it is on the edit-profile page.
// It is used in this if statement so that it doesn't rerun when the HTML page changes.
if (window.location.href.split('/').pop() != 'edit-profile.html') {
    window.api.getCurrentCompany().then(user => {
        console.log(user);

        // Handles creating a new company
        const editInfo = document.getElementById('edit-info');
        console.log('middle', editInfo);

        const createCompanyButton = document.getElementById('create-company');
        console.log('cc-button', createCompanyButton);
        createCompanyButton.addEventListener('click', e => {
            // e.preventDefault();
            const createCompanyTemplate = document.getElementById('create-company-template');
            const createCompanyForm = createCompanyTemplate.content.cloneNode(true);
            editInfo.innerText = "";
            editInfo.appendChild(createCompanyForm);

            const submitButton = document.getElementById('create-company-submit');
            submitButton.addEventListener('click', e => {
                e.preventDefault();
                const name = document.getElementById('company-name').value;
                const number = document.getElementById('invoice-number').value;
                const address = document.getElementById('company-address').value;
                const city = document.getElementById('company-city').value;
                const zipCode = document.getElementById('company-zip-code').value;

                const response = window.api.createCompany({name, number, address, city, zipCode});
                console.log('create res:', response);
            })
        })

        // Handles populating companies on side
        const companyList = document.getElementById('company-list');
        console.log('list', companyList);
        window.api.getCompanies().then(companies => {
            console.log('companies', companies);
            companies.forEach(company => {
                console.log('c', company);
                const companyProfileTemplate = document.getElementById('company-template');
                const companyProfile = companyProfileTemplate.content.cloneNode(true);
                companyProfile.getElementById('company-profile-name').innerText = company.name;
                companyProfile.getElementById('company-profile-name').id = company.id;
                companyList.prepend(companyProfile);

                // Adds click event listener to button on side
                const currentCompanyProfile = document.getElementById(company.id);
                console.log('company profile', currentCompanyProfile);
                currentCompanyProfile.addEventListener('click', e => {
                    const companyFormTemplate = document.getElementById('edit-company-template');
                    const companyForm = companyFormTemplate.content.cloneNode(true);
                    companyForm.getElementById('company-name').value = company.name;
                    companyForm.getElementById('invoice-number').value = company.invoice_number;
                    companyForm.getElementById('company-address').value = company.address;
                    companyForm.getElementById('company-city').value = company.city;
                    companyForm.getElementById('company-zip-code').value = company.zip_code;
                    companyForm.getElementById('company-quantity').value = company.quantity;
                    companyForm.getElementById('company-unit-price').value = company.unit_price;
                    companyForm.getElementById('company-description').value = company.description;

                    editInfo.innerText = "";
                    editInfo.appendChild(companyForm);

                    const saveButton = document.getElementById('save');
                    console.log('save', saveButton);
                    saveButton.addEventListener('click', e => {
                        const name = document.getElementById('company-name').value;
                        const invoiceNumber = document.getElementById('invoice-number').value;
                        const address = document.getElementById('company-address').value;
                        const city = document.getElementById('company-city').value;
                        const zipCode = document.getElementById('company-zip-code').value;
                        const quantity = document.getElementById('company-quantity').value;
                        const unitPrice = document.getElementById('company-unit-price').value;
                        const description = document.getElementById('company-description').value;
                        const id = company.id;
                        const response = window.api.updateCompany({name, invoiceNumber, address, city, zipCode, quantity, unitPrice, description, id})
                        console.log('update res:', response);
                    });
                });
            })
        }).catch(err => {
            console.log('err', err);
        })

    }).catch(err => {
        // window.api.createCompany({name, number, address, city, zipCode});
        window.location.href = 'edit-profile.html';
    })
} else {
    // will be run once on the edit-profile page
    const submitButton = document.getElementById('create-initial-company-submit');
    console.log('sb', submitButton);
    submitButton.addEventListener('click', e => {
        e.preventDefault();
        const name = document.getElementById('company-name').value;
        const address = document.getElementById('company-address').value;
        const city = document.getElementById('company-city').value;
        const zipCode = document.getElementById('company-zip-code').value;
        const phoneNumber = document.getElementById('company-phone-number').value;
        const response = window.api.createInitialCompany({name, address, city, zipCode, phoneNumber});
        console.log('res:', response);
        window.location.href = 'index.html';
    })
}
// window.location.href = 'edit-profile.html';

