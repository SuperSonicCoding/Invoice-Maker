// Handles initial startup
console.log('window', window.location.href.split('/').pop());
// Gets the last thing in the file location and uses it to see if it is on the edit-profile page.
// It is used in this if statement so that it doesn't rerun when the HTML page changes.
if (window.location.href.split('/').pop() != 'edit-profile.html') {
    window.api.getCurrentCompany().then(user => {
        console.log(user);

        // Handles populating companies on side
        const companyList = document.getElementById('company-list');
        console.log('list', companyList);
        window.api.getCompanies().then(companies => {
            console.log('companies', companies);
            companies.forEach(company => {
                console.log('c', company);
                const companyProfileTemplate = document.getElementById('company-template');
                const companyProfile = companyProfileTemplate.content.cloneNode(true);
                companyProfile.getElementById('company-profile-button').innerText = company.name;
                companyProfile.getElementById('company-profile-button').id = company.id;
                companyList.prepend(companyProfile);

                // Adds click event listener to button on side
                const currentCompanyProfile = document.getElementById(company.id);
                console.log('company profile', currentCompanyProfile);
                currentCompanyProfile.addEventListener('click', e => {
                    const companyProfiles = document.querySelectorAll('.company');
                    companyProfiles.forEach(companyProfile => {
                        companyProfile.classList.remove('selected');
                    });
                    currentCompanyProfile.classList.add('selected');
                    const companyFormTemplate = document.getElementById('edit-company-template');
                    const companyForm = companyFormTemplate.content.cloneNode(true);
                    companyForm.getElementById('company-name').value = company.name;
                    companyForm.getElementById('invoice-number').value = company.invoice_number;

                    // getting current date from stack overflow
                    // https://stackoverflow.com/questions/23593052/format-javascript-date-as-yyyy-mm-dd
                    let today = new Date();
                    const offset = today.getTimezoneOffset();
                    today = new Date(today.getTime() - (offset * 60 * 1000));
                    companyForm.getElementById('date').value = today.toISOString().split('T')[0];
                    // end of stack overflow usage

                    companyForm.getElementById('company-address').value = company.address;
                    companyForm.getElementById('company-city').value = company.city;
                    companyForm.getElementById('state-initials').value = company.state_initials;
                    companyForm.getElementById('company-zip-code').value = company.zip_code;
                    companyForm.getElementById('company-quantity').value = company.quantity;
                    companyForm.getElementById('company-unit-price').value = company.unit_price;
                    companyForm.getElementById('company-description').value = company.description;

                    editInfo.innerText = "";
                    editInfo.appendChild(companyForm);

                    const editCompanyForm = document.getElementById('edit-company-form');
                    editCompanyForm.addEventListener('submit', e => {
                        e.preventDefault();
                        const buttonId = e.submitter.id;
                        if (buttonId == 'save') {
                            console.log('save save save');
                            const name = document.getElementById('company-name').value;
                            const invoiceNumber = document.getElementById('invoice-number').value;
                            const address = document.getElementById('company-address').value;
                            const city = document.getElementById('company-city').value;
                            const stateInitials = document.getElementById('state-initials').value;
                            const zipCode = document.getElementById('company-zip-code').value;
                            const quantity = document.getElementById('company-quantity').value;
                            const unitPrice = document.getElementById('company-unit-price').value;
                            const description = document.getElementById('company-description').value;
                            const id = company.id;
                            const response = window.api.updateCompany({name, invoiceNumber, address, city, stateInitials, zipCode, quantity, unitPrice, description, id})
                            console.log('update res:', response);
                        } else if (buttonId == 'create-file-button') {
                            const fullName = user.full_name;
                            const currentCompanyName = user.name;
                            const currentCompanyAddress = user.address;
                            const currentCompanyCity = user.city;
                            const currentCompanyStateInitials = user.state_initials;
                            const currentCompanyZipCode = user.zip_code;
                            const phoneNumber = user.phone_number;
                            const email = user.email;

                            // const companyProfileName = document.getElementById('company-name').value;
                            // const invoiceNumber = document.getElementById('invoice-number').value;
                            // const date = document.getElementById('date').value;
                            // const companyProfileAddress = document.getElementById('company-address').value;
                            // const companyProfileCity = document.getElementById('company-city').value;
                            // const companyProfileStateInitials = document.getElementById('state-initials').value;
                            // const companyProfileZipCode = document.getElementById('company-zip-code').value;
                            // const quantity = document.getElementById('company-quantity').value;
                            // const unitPrice = document.getElementById('company-unit-price').value;
                            // const description = document.getElementById('company-description').value;

                            const companyProfileName = company.name;
                            const invoiceNumber = company.invoice_number;
                            const date = document.getElementById('date').value;
                            const companyProfileAddress = company.address;
                            const companyProfileCity = company.city;
                            const companyProfileStateInitials = company.state_initials;
                            const companyProfileZipCode = company.zip_code
                            const quantity = company.quantity;
                            const unitPrice = company.unit_price;
                            const description = company.description;
                            const filePath = company.file_path;
                            const id = company.id;
                            const response = window.api.createFile({fullName, currentCompanyName, currentCompanyAddress, currentCompanyCity, currentCompanyStateInitials,
                                currentCompanyZipCode, phoneNumber, email, companyProfileName, invoiceNumber, date, companyProfileAddress, companyProfileCity, companyProfileStateInitials, companyProfileZipCode, quantity, unitPrice, description, filePath, id});
                            console.log('file creation res', response);
                        }
                    })

                });
            })
        }).catch(err => {
            console.log('err', err);
        })

        // Handles creating a new company
        const editInfo = document.getElementById('edit-info');
        console.log('middle', editInfo);

        const createCompanyButton = document.getElementById('create-company-button');
        console.log('cc-button', createCompanyButton);
        createCompanyButton.addEventListener('click', e => {
            // e.preventDefault();
            const companyProfiles = document.querySelectorAll('.company');
            console.log('profiles', companyProfiles);
            companyProfiles.forEach(companyProfile => {
                companyProfile.classList.remove('selected');
            });
            createCompanyButton.classList.add('selected');
            const createCompanyTemplate = document.getElementById('create-company-template');
            const createCompany = createCompanyTemplate.content.cloneNode(true);
            editInfo.innerText = "";
            editInfo.appendChild(createCompany);

            const createCompanyForm = document.getElementById('create-company-form');
            createCompanyForm.addEventListener('submit', e => {
                e.preventDefault();
                const name = document.getElementById('company-name').value;
                const number = document.getElementById('invoice-number').value;
                const address = document.getElementById('company-address').value;
                const city = document.getElementById('company-city').value;
                const stateInitials = document.getElementById('state-initials').value;
                const zipCode = document.getElementById('company-zip-code').value;

                const response = window.api.createCompany({name, number, address, city, stateInitials, zipCode});
                console.log('create res:', response);
                location.reload();
            })

            // const submitButton = document.getElementById('create-company-submit');
            // submitButton.addEventListener('submit', e => {
            //     e.preventDefault();
            //     const name = document.getElementById('company-name').value;
            //     const number = document.getElementById('invoice-number').value;
            //     const address = document.getElementById('company-address').value;
            //     const city = document.getElementById('company-city').value;
            //     const stateInitials = document.getElementById('state-initials').value;
            //     const zipCode = document.getElementById('company-zip-code').value;

            //     const response = window.api.createCompany({name, number, address, city, stateInitials, zipCode});
            //     console.log('create res:', response);
            // })
        })

        // Handles if the user wants to edit their profile
        const editProfileButton = document.getElementById('edit-profile');
        console.log('edit', editProfileButton);
        editProfileButton.addEventListener('click', event => {
            window.location.href = 'edit-profile.html';
        });
    }).catch(err => {
        // window.api.createCompany({name, number, address, city, zipCode});
        window.location.href = 'edit-profile.html';
    })
} else {
    // will be run once on the edit-profile page
    const currentCompanyForm = document.getElementById('create-company-form');
    
    // handles phone number input, allowing only numbers and adding in formatting
    const phoneNumberInput = document.getElementById('company-phone-number');
    phoneNumberInput.addEventListener('keydown', e => {
        const keyPressed = e.key;
        const keyIsDigit = /^\d$/.test(keyPressed); // sees if the input is a digit or not
        const allowedKeys = ['Backspace']
        if ((!keyIsDigit && !allowedKeys.includes(keyPressed)) || keyPressed == ' ') {
            e.preventDefault()
        }

        // adds formatting when typing digits
        if (keyIsDigit && phoneNumberInput.value.length == 0) {
            e.preventDefault()
            phoneNumberInput.value += `(${keyPressed}`;
        }

        if (keyIsDigit && phoneNumberInput.value.length == 3) {
            e.preventDefault()
            phoneNumberInput.value += `${keyPressed}) `;
        }

        if (keyIsDigit && phoneNumberInput.value.length == 9) {
            e.preventDefault()
            phoneNumberInput.value += `-${keyPressed}`;
        }

        // removes formatting when backspace
        if (keyPressed == 'Backspace') {
            if (phoneNumberInput.value.length == 11) {
                phoneNumberInput.value = phoneNumberInput.value.slice(0, -1); // need to remember slice returns the string
            }

            if (phoneNumberInput.value.length == 7) {
                phoneNumberInput.value = phoneNumberInput.value.slice(0, -2); // need to remember slice returns the string
            }

            if (phoneNumberInput.value.length == 2) {
                phoneNumberInput.value = phoneNumberInput.value.slice(0, -1); // need to remember slice returns the string
            }
        }
    });

    window.api.getCurrentCompany().then(user => {
        // adds a cancel button to go back to the home page
        const cancelButton = document.createElement('button');
        cancelButton.id = 'cancel-edit-company';
        cancelButton.classList.add('my-2', 'w-100');
        cancelButton.innerText = 'Cancel';
        cancelButton.addEventListener('click', e => {
            window.location.href = 'index.html';
        });
        currentCompanyForm.appendChild(cancelButton);

        document.getElementById('title').innerText = "Update Company Profile Data";
        // populates fields
        document.getElementById('company-name').value = user.name;
        document.getElementById('company-address').value = user.address;
        document.getElementById('company-city').value = user.city;
        document.getElementById('state-initials').value = user.state_initials;
        document.getElementById('company-zip-code').value = user.zip_code;
        document.getElementById('company-phone-number').value = user.phone_number;
        document.getElementById('company-email').value = user.email;

        currentCompanyForm.addEventListener('submit', e => {
            e.preventDefault();
            const fullName = document.getElementById('full-name').value;
            const companyName = document.getElementById('company-name').value;
            const address = document.getElementById('company-address').value;
            const city = document.getElementById('company-city').value;
            const stateInitials = document.getElementById('state-initials').value
            const zipCode = document.getElementById('company-zip-code').value;
            const phoneNumber = document.getElementById('company-phone-number').value;
            const email = document.getElementById('company-email').value;
            const response = window.api.updateCurrentCompany({fullName, companyName, address, city, stateInitials, zipCode, phoneNumber, email});
            console.log('res:', response);
            window.location.href = 'index.html';
        })
    }).catch(err => {
        document.getElementById('title').innerText = "Create Company Profile Data";
        currentCompanyForm.addEventListener('submit', e => {
            e.preventDefault();
            const fullName = document.getElementById('full-name').value;
            const companyName = document.getElementById('company-name').value;
            const address = document.getElementById('company-address').value;
            const city = document.getElementById('company-city').value;
            const stateInitials = document.getElementById('state-initials').value
            const zipCode = document.getElementById('company-zip-code').value;
            const phoneNumber = document.getElementById('company-phone-number').value;
            const response = window.api.createInitialCompany({fullName, companyName, address, city, zipCode, stateInitials, phoneNumber});
            console.log('res:', response);
            window.location.href = 'index.html';
        })
    });
}
// window.location.href = 'edit-profile.html';

