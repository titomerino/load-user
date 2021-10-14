import xlrd
from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group

teacher_group = Group.objects.get(name='Teachers')
UserModel = get_user_model()

filePath = "teachers.xlsx"
openFile = xlrd.open_workbook(filePath)
sheet = openFile.sheet_by_name("202122")
contExist = 0
contNotExist = 0
for i in range(sheet.nrows):
    if i != 0 and sheet.cell_value(i, 0) != "":
        username = sheet.cell_value(i, 0).lower() + "-" + sheet.cell_value(i, 1).lower()
        email = sheet.cell_value(i, 4).lower()

        if (UserModel.objects.filter(email=email).exists()):
            contExist = contExist + 1
        else:
            if (UserModel.objects.filter(username=username).exists()):
                username = username + '-s'
            # Create user
            me = UserModel.objects.create(
                username=username,
                first_name=sheet.cell_value(i, 0),
                last_name=sheet.cell_value(i, 1),
                email=sheet.cell_value(i, 4).lower(),
            )
            me.set_password('123456789')
            try:
                me.save()
                # insert user to group
                teacher_group.user_set.add(me)
                print('Added user: ', me.username, '', me.email)
                contNotExist = contNotExist + 1
            except Exception:
                print('User not added: ', sheet.cell_value(i, 0), '', email)

print('Total not added: ', contExist)
print('Total added: ', contNotExist)
