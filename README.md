# ControlId

Exemplo de utilização do SDK de integração RepCid.dll da ControlId para comunicação do relógio de ponto Rep IdClass com aplicações criadas com o Visual Basic 6 ou 5.

Este é um exemplo desenvolvido pela empresa __Liondas Softwares__, produtora do programa para tratamento do cartão de ponto __LdsPonto__.
 
[www.liondas.com.br](http://www.liondas.com.br)

###Pré Requisitos

---

- SDK RepCid.dll vrs.7.11 instalado e registrado para ser utilizado como interface COM.

- Visual Basic 5 ou 6.

- Este exemplo foi testado no Windows 10 64 bits. 

- Relógio de ponto modelo Rep IdClass.

###Download do SDK

---

O SDK deverá ser solicitado para o fabricante ControlId, no site [www.controlid.com.br](http://www.controlid.com.br).

###Considerações

---

Neste exemplo, além das funções de conexão, foi implementado o controle de erros, que ao meu ver é de grande importância.

Para uma melhor leitura do exemplo, foi dado prioridade para um código limpo e simples, utilizando os nomes padrões dos objetos (form, buttons, labels)

Uma dica, a _senha Web_ do equipamento deverá estar no padrão de fábrica, para isso consulte o manual do equipamento.

Os nomes ControlId, Id Class e RepCid são marcas da empresa ControlId.
