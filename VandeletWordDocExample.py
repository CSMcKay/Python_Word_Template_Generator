# -*- coding: utf-8 -*-
'''
Created : 2016-03-26

@author: Eric Lapouyade
'''

from docxtpl import DocxTemplate

#tpl=DocxTemplate('marketing_doc_template_binder_tpl.docx')
tpl=DocxTemplate('sage_mk_tpl.docx')

context = {
'current_date' : '31 October 2017',
    'client_name' : 'Art E. Vandelet',
    'client_address_street' : '544 DogGone, Dr.',
    'client_address_city' : 'Corinth',
    'client_address_state' : 'MS',
    'client_address_zip' : '38834',
    'client_phone_number' : '(931) 412-7120',
    'client_social_security_number' : '111-79-812',
    'client_ownership_percentage' : '4',

    'sh1_name' : 'Ed Goldein',
    'sh1_address_street' : 'sh1_address_street',
    'sh1_address_city' : 'sh1_address_city',
    'sh1_address_state' : 'sh1_address_state',
    'sh1_address_zip' : 'sh1_address_zip',
    'sh1_phone_number' : 'sh1_phone_number',
    'sh1_social_security_number' : 'sh1_social_security_number',
    'sh1_ownership_percentage' : '48',

    'sh2_name': 'Mary Love',
    'sh2_address_street': 'sh1_address_street',
    'sh2_address_city': 'sh1_address_city',
    'sh2_address_state': 'sh1_address_state',
    'sh2_address_zip': 'sh1_address_zip',
    'sh2_phone_number': 'sh1_phone_number',
    'sh2_social_security_number': 'sh1_social_security_number',
    'sh2_ownership_percentage': '48',

    'operating_company_name' : 'Pizza Urgent Care of Logan, PA',
    'operating_company_state_of_organization' : 'Mississippi',
    'operating_company_address_street' : '21 Geell Road, Suite 12',
    'operating_company_address_city' : 'Logan',
    'operating_company_address_state' : 'MS',
    'operating_company_address_zip' : '38234',
    'operating_company_ein' : '64-0232721',
    'operating_company_manager_name' : 'Art E. Vandelet',
    'operating_company_contact_name' : 'Art E. Vandelet',
    'operating_company_contact_title' : 'Manager',
    'operating_company_contact_phone' : '(662) 415-7019',

    'marketing_company_name': 'PET Ambulance, Inc.',
    'marketing_company_state_of_organization': 'Mississippi',
    'marketing_company_address_street': '14208 Pline Road',
    'marketing_company_address_city': 'Paris',
    'marketing_company_address_state': 'MS',
    'marketing_company_address_zip': '38834',
    'marketing_company_ein': '47-423873',
    'marketing_company_president_name': 'marketing_company_president_name',
    'marketing_company_manager_name': 'marketing_company_manager_name',
    'marketing_company_contact_name': 'Art E. Vandelet',
    'marketing_company_contact_title': 'marketing_company_contact_title',
    'marketing_company_contact_phone': '(634) 423-7880',

    'family_holding_company_name': 'Vandelet Enterprises, LP',
    'family_holding_company_entity_type' : 'Limited Partnership',
    'family_holding_company_state_of_organization' : 'Mississippi',
    'family_holding_company_address_street': '263 ARr Road, Suite 1',
    'family_holding_company_address_city': 'Corinth',
    'family_holding_company_address_state': 'MS',
    'family_holding_company_address_zip': '38834',
    'family_holding_company_ein': '64-0895593',
    'family_holding_company_manager_name': 'Art E. Vandelet',
    'family_holding_company_contact_name': 'family_holding_company_contact_name',
    'family_holding_company_contact_title': 'Manager',
    'family_holding_company_phone': 'family_holding_company_phone',

    'new_management_company_name': 'new_management_company_name',
    'new_management_company_entity_type': 'new_management_company_entity_type',
    'new_management_company_state_of_organization': 'new_management_company_state_of_organization',
    'new_management_company_address_street': 'new_management_company_address_street',
    'new_management_company_address_city': 'new_management_company_address_city',
    'new_management_company_address_state': 'MS',
    'new_management_company_address_zip': 'new_management_company_address_zip',
    'new_management_company_ein': 'new_management_company_ein',
    'new_management_company_manager_name': 'new_management_company_manager_name',
    'new_management_company_contact_name': 'new_management_company_contact_name',
    'new_management_company_contact_title': 'new_management_company_contact_title',
    'new_management_company_phone': 'new_management_company_phone',

    'attorney_name': 'Ross Dixon, PLLC',
    'attorney_address_street': '40 S. 111th St., Ste. 1014',
    'attorney_address_city': 'Corinth',
    'attorney_address_state': 'MS',
    'attorney_address_zip': '38834',
    'attorney_phone_number': 'Office: (122) 312-4112',

    'date': '2016-03-17',

}

import os
directory='WORD_DOCS_OUTPUT/Vandelet/'
if not os.path.exists(directory):
    os.makedirs(directory)
tpl.render(context)
tpl.save(directory+'Vandelet_binder.docx')