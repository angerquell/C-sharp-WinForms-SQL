SELECT r.��� as "����� �����������",
COUNT(*) As "���������� �����������  ���������"
FROM [�����] r JOIN [�������] v
on r.���_������= v.���_������
GROUP BY r.���
ORDER BY "���������� �����������  ���������" DESC;
