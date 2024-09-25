Este código foi utilizado internamente. 

Um Servidor SQL foi hospedado em uma maquina em nossa sala, este servidor coletava dados de diferentes locais 
da empresa e convertir em arquivo xlsx para ser usado com gráfico Power B.I.

Infelizmente não foi possível anexar o gráfico B.I pois é uma métrica passou a ser usada pela empresa para avaliação, 
na qual não foi permitida a utilização para divulgação em estudos.

O código possuí disponibilizado em seus sistema cadastro de dados manuais, pois nem todos os dados necessarios possuiam banco de dados. Então foi gerada uma ferramenta de 
cadastro de dados, na qual a cada lançamento, era alterado os dados de equipamento de laboratório disponibilizados.

Foi adaptado um banco de dados SQLite para hospedar alguns dados basicos para que a ferramenta possa ser utilizada sem conexão com o sistema interno da empresa.

Os dados de conexão foram removidos do código afim de mante-lo o mais limpo possível.
