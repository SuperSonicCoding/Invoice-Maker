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
        console.log('res:', response);
    })
})


// Handles initial startup
// window.api.getCurrentCompany().then(user => {
//     console.log(user);
// }).catch(err => {
//     const name = "";
//     const number = "";
//     const address = "";
//     const city = "";
//     const zipCode = "";
//     window.api.createCompany({name, number, address, city, zipCode});
//     window.location.href = 'edit-profile.html';
// })
window.location.href = 'edit-profile.html';

// Handles populating companies on side

