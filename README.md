# Estrela Geométrica
Estrela criada por retas - em VBA

Como é época de Natal, a seguir um código para desenhar uma estrela só utilizando segmentos de reta e geometria, para alegrar as festas.

Imagine duas retas perpendiculares.

![](https://ferramentasexcelvba.files.wordpress.com/2019/12/estrela01.png)

Cada reta recebe N divisões.

 

Agora, uma o primeiro ponto do eixo X ao ponto logo acima do meio do eixo Y, o segundo ponto do eixo X ao próximo ponto acima do eixo Y, e repita até o fim.

O comando VBA para adicionar uma linha é, basicamente:
ActiveSheet.Shapes.AddLine(x0, y0, x1, y1)

E tem outros comandos para colorir, informar a espessura e controlar. Mas, basicamente, o único comando especial é esse de criar retas.



O resultado não impressiona muito. Porém, se aumentarmos o número de pontos, fica mais divertido.
![](https://ferramentasexcelvba.files.wordpress.com/2019/12/estrela02.png)

 

Com 10 pontos:
![](https://ferramentasexcelvba.files.wordpress.com/2019/12/estrela03.png)

Com 20 pontos:
![](https://ferramentasexcelvba.files.wordpress.com/2019/12/estrela04.png)
 

Com 35 pontos:
![](https://ferramentasexcelvba.files.wordpress.com/2019/12/estrela05.png)
 

Não sei exatamente o nome desta curva (Marcos Melo, sabe?), porém, a técnica é muito conhecida na matemática.


