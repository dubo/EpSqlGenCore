select  a.first_name, a.last_name , f.title, f.release_year, f.rating  
 from actor a , film f , film_actor fa 
 where a.last_name = :lastName   --'Davis' 
   and a.actor_id = fa.actor_id
   and fa.film_id =  f.film_id  
 order by f.rating