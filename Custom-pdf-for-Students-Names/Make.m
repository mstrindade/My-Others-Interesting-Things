% Set your path:
BANCODADOS = 'students.xlsx';
PLANILHA = 'GAT'; PROVACORPO = 'provacorpo.tex';

[ndata, text, alldata] = xlsread(BANCODADOS,PLANILHA);

[m, n] = size(alldata);

done = 0; while done == 0; try                                  %#ok<ALIGN>
for nome = 2:m
    ALUNO = alldata{nome,4};
    CURSO = alldata{nome,1};
    MATRICULA = alldata{nome,3};
    fprintf(['Lido: ',ALUNO,' (',num2str(MATRICULA),') \n']);

    % ABRE A PROVACORPO TROCA O NOME E SALVA NOVO 23572864752345.TEX
    fid = fopen(PROVACORPO,'rt');
    X = fread(fid);
    fclose(fid);
    X = char(X.');
    Y = strrep(X, 'NOMEDOGAJO', [ALUNO,' (',num2str(MATRICULA),') ']);
    Y = strrep(Y, 'CURSODOGAJO', CURSO);
    
    fid2 = fopen([num2str(MATRICULA),'.tex'],'wt');
    fwrite(fid2,Y);
    fclose(fid2);
    
    
    % ABRE ARQUIVO UNICO E INSERE CADA PROVA
    if nome == 2
        fid = fopen('prova.tex','rt');
        delete('provaFULL.tex');
    else
        fid = fopen('provaFULL.tex','rt');
    end
    X = fread(fid);
    fclose(fid);
    X = char(X.');
    Y = strrep(X, 'INSERTHERE', ['\input{',num2str(MATRICULA),'}INSERTHERE']);
    
    fprintf(['insreindo (',num2str(MATRICULA),') \n']);
    
    fid2 = fopen('provaFULL.tex','wt');
    fwrite(fid2,Y);
    fclose(fid2);
    
end
    fid = fopen('provaFULL.tex','rt');
    X = fread(fid);
    fclose(fid);
    X = char(X.');
    Y = strrep(X, 'INSERTHERE', ' ');
    
    fid2 = fopen('provaFULL.tex','wt');
    fwrite(fid2,Y);
    fclose(fid2);
    
done = 1; catch;  done = 0; end; end