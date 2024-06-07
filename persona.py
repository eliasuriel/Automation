class Persona:
    def __init__(self, nombre: str):
        self._nombre = nombre
        self._proyectos = []

    def nombre(self):
        return self._nombre

    def proyectos(self):
        return self._proyectos
    
    
    def tiene_proyecto(self, nombre_proyecto, mes=None):
        for proyecto in self._proyectos:
            if nombre_proyecto in proyecto['nombre_proyecto']  and (mes is None or proyecto['mes'] == mes):
                return True
        return False
    
    def tiene_componente_en_proyecto(self, nombre_proyecto, nombre_componente, mes=None):
        for proyecto in self._proyectos:
            if nombre_proyecto in proyecto['nombre_proyecto'] and (mes is None or proyecto['mes'] == mes):
                for componente in proyecto['componentes']:
                    if componente['nombre_componente'] == nombre_componente:
                        return True
        return False

    def actualizar_horas_componente(self, nombre_proyecto: str, nombre_componente: str, horas_extra: int, mes: str):
        for proyecto in self._proyectos:
            if nombre_proyecto in proyecto['nombre_proyecto'] and proyecto['mes'] == mes:
                for componente in proyecto['componentes']:
                    if componente['nombre_componente'] == nombre_componente:
                        componente['horas'] += horas_extra
                        break
                break

    def agregar_proyecto_simple(self, nombre_proyecto: str, mes: str, fuente: str):
        if not self.tiene_proyecto(nombre_proyecto, mes):
            proyecto = {
                'nombre_proyecto': nombre_proyecto,
                'componentes': [],
                'mes': mes,
                'fuente': fuente
            }
            self._proyectos.append(proyecto)
    
    def get_fuente_proyecto(self, nombre_proyecto: str, mes: str = None):
        for proyecto in self._proyectos:
            if (nombre_proyecto in proyecto['nombre_proyecto']) and (mes in proyecto['mes'] ):
                return True
        return False
    
    def get_nombres_proyectos(self):
        return [proyecto['nombre_proyecto'] for proyecto in self._proyectos]
       

    def agregar_componente_a_proyecto(self, nombre_proyecto: str, nombre_componente: str, horas: int, mes: str):
        for proyecto in self._proyectos:
            if nombre_proyecto in proyecto['nombre_proyecto'] and proyecto['mes'] == mes:
                if not self.tiene_componente_en_proyecto(nombre_proyecto, nombre_componente, mes):
                    nuevo_componente = {'nombre_componente': nombre_componente, 'horas': horas}
                    proyecto['componentes'].append(nuevo_componente)
                else:
                    self.actualizar_horas_componente(nombre_proyecto, nombre_componente, horas, mes)
                    #print(self._nombre)
                    #print(f"Se agregaran estas horas '{horas}' al componente '{nombre_componente}' del proyecto '{nombre_proyecto}' en el mes '{mes}'. \n")
                break
        else:
            print(f"El proyecto '{nombre_proyecto}' no existe en el mes '{mes}'.")

    def agregar_horas_a_componente(self, nombre_proyecto: str, nombre_componente: str, horas: int, mes: str):
        for proyecto in self._proyectos:
            if  nombre_proyecto in proyecto['nombre_proyecto'] and proyecto['mes'] == mes:
                if not self.tiene_componente_en_proyecto(nombre_proyecto, nombre_componente, mes):
                    nuevo_componente = {'nombre_componente': nombre_componente, 'horas': horas}
                    proyecto['componentes'].append(nuevo_componente)
                else:
                    self.set_horas_componente(nombre_proyecto, nombre_componente, horas, mes)
                    #print(self._nombre)
                    #print(f"Se agregaran estas horas '{horas}' al componente '{nombre_componente}' del proyecto '{nombre_proyecto}' en el mes '{mes}'. \n")
                break
        else:
            print(f"El proyecto '{nombre_proyecto}' no existe en el mes '{mes}'.")

    def get_componentes(self, nombre_proyecto: str, mes: str):
        for proyecto in self._proyectos:
            if nombre_proyecto in proyecto['nombre_proyecto'] and proyecto['mes'] == mes:
                return proyecto['componentes']
            
            
        

    def set_componentes(self, nombre_proyecto: str, componentes: list, mes: str):
        for proyecto in self._proyectos:
            if nombre_proyecto in proyecto['nombre_proyecto'] and proyecto['mes'] == mes:
                proyecto['componentes'] = componentes
                break

    def get_horas_componente(self, nombre_proyecto: str, nombre_componente: str, mes: str):
        for proyecto in self._proyectos:
            if nombre_proyecto in proyecto['nombre_proyecto'] and proyecto['mes'] == mes:
                for componente in proyecto['componentes']:
                    if componente['nombre_componente'] == nombre_componente:
                        return componente['horas']
        return 0

    def set_horas_componente(self, nombre_proyecto: str, nombre_componente: str, horas: int, mes: str):
        for proyecto in self._proyectos:
            if nombre_proyecto in proyecto['nombre_proyecto']  and proyecto['mes'] == mes:
                for componente in proyecto['componentes']:
                    if componente['nombre_componente'] == nombre_componente:
                        componente['horas'] = horas
                        break

    def get_proyecto(self, nombre_proyecto: str, mes: str = None):
        proyectos = []
        for proyecto in self._proyectos:
            if nombre_proyecto in proyecto['nombre_proyecto'] and (mes is None or proyecto['mes'] == mes):
                proyectos.append(proyecto)
        return proyectos if proyectos else None
    
    def get_nombres_proyectos_por_fuente(self, fuente: str):
        return [proyecto['nombre_proyecto'] for proyecto in self._proyectos if proyecto['fuente'] == fuente]
    
    def get_meses_proyecto(self, nombre_proyecto: str):
        meses = []
        for proyecto in self._proyectos:
            if nombre_proyecto in proyecto['nombre_proyecto']:
                meses.append(proyecto['mes'])
        return meses

    def __str__(self):
        proyectos_str = ""
        for proyecto in self._proyectos:
            componentes_str = "\n  ".join([f"Componente: {c['nombre_componente']}, Horas: {c['horas']}" for c in proyecto['componentes']])
            proyectos_str += f"Proyecto: {proyecto['nombre_proyecto']}, Mes: {proyecto['mes']}, Fuente: {proyecto['fuente']}\n  {componentes_str}\n"
        return f"Nombre: {self._nombre}\n{proyectos_str}"

    def detalles(self):
        return str(self)
    
