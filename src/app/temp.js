db.createCollection("nonfiction", {
    validator:{
        $jsonSchema: {
            required:['name','price'],
            properties:{
                name:{
                    bsonType:'string',
                    description: 'must be a string and required'
                },price:{
                    bsonType:'number',
                    description:'must be a number and required'
                }
            }
        }

    },
    validationAction:'error'
    // validation:'' //it's a string if warn,then in log folder we see warn,if validation failed or if we write error,it shows error if failed and won't let it insert,by default it's error
})