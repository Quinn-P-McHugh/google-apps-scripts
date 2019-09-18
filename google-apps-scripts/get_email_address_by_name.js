/* The following document uses JSDoc 3 standard for JavaScript documentation: http://usejsdoc.org/ */

/**
 * Searchs for a person in your contacts by name and returns
 * the email address of the first search result.
 *
 * @param {String} name The full name of the desired person.
 */
function getEmailAddressByName(name) {
  var contact = ContactsApp.getContactsByName(name)[0];
  var email = contact.getPrimaryEmail();
  return email;
}
