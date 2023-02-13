# Outlook-Envio-Email-Automatico
Esse código lê uma tabela em excel e utiliza os dados da tabela para enviar um email personalizado conforme os dados da tabela

O código abaixo é utilizado em ambiente de produção na minha empresa
Usamos ele após outras rotinas de bancos de dados cruzar oportunidades de acesso em imóveis de difícil acesso pelas equipes técnicas
Quando identifica uma oportunidade, popula uma tabela que atualiza de hora em hora

O nosso código python percorre essa tabela também de hora em hora, e caso identifique um caso da cidade do Rio de Janeiro, realiza o envio de um email para a equipe
de gestores responsável pela região sinalizando a oportunidade de acesso, acompanhado dos dados do serviço identificado.

Pode ser facilmente escalado para qualquer tabela excel, é só substituir as variáveis pelos novos campos de colunas e personalizar a mensagem e mailing de acordo com a necessidade.
